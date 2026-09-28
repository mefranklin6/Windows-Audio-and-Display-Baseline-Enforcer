import 'dart:convert';
import 'dart:io';
import 'dart:math';

const audioConfigurationFileNames = <String>[
  'audio_device_list.json',
  'audio_levels.json',
];
const displayConfigurationFileName = 'display_config_profile.xml';
const configurationFileNames = <String>[
  ...audioConfigurationFileNames,
  displayConfigurationFileName,
];

enum BackupConflictPolicy { overwriteAll, missingPcsOnly }

enum RestoreConflictPolicy { overwriteAll, skipAllExisting }

class BackupVersion {
  const BackupVersion({
    required this.pc,
    required this.folder,
    required this.createdAt,
    required this.isLatest,
  });

  final String pc;
  final Directory folder;
  final DateTime createdAt;
  final bool isLatest;

  String get label {
    final local = createdAt.toLocal();
    final stamp =
        '${local.year.toString().padLeft(4, '0')}-'
        '${local.month.toString().padLeft(2, '0')}-'
        '${local.day.toString().padLeft(2, '0')} '
        '${local.hour.toString().padLeft(2, '0')}:'
        '${local.minute.toString().padLeft(2, '0')}:'
        '${local.second.toString().padLeft(2, '0')}';
    return isLatest ? 'Latest — $stamp' : 'Historical — $stamp';
  }
}

class TargetConfigurationState {
  const TargetConfigurationState({
    required this.files,
    required this.modifiedAt,
  });

  final Map<String, bool> files;
  final Map<String, DateTime> modifiedAt;

  bool get hasAny => files.values.any((exists) => exists);
  bool get hasAll => files.values.every((exists) => exists);
}

class ConfigurationOperationResult {
  const ConfigurationOperationResult({
    required this.pc,
    required this.success,
    required this.copiedFiles,
    this.error = '',
  });

  final String pc;
  final bool success;
  final List<String> copiedFiles;
  final String error;
}

class ConfigurationBackupRepository {
  const ConfigurationBackupRepository();

  String safePcName(String pc) {
    final safe = pc
        .trim()
        .replaceAll(RegExp(r'[^A-Za-z0-9_.-]+'), '_')
        .replaceAll(RegExp(r'^[._]+|[._]+$'), '');
    return safe.isEmpty ? 'unknown' : safe.toLowerCase();
  }

  Directory pcFolder(String root, String pc) =>
      Directory(_join(root, safePcName(pc)));

  Directory latestFolder(String root, String pc) =>
      Directory(_join(pcFolder(root, pc).path, 'latest'));

  Future<bool> hasBackup(String root, String pc) async {
    final latest = latestFolder(root, pc);
    return configurationFileNames.any(
      (name) => File(_join(latest.path, name)).existsSync(),
    );
  }

  Future<List<BackupVersion>> versions(String root, String pc) async {
    final result = <BackupVersion>[];
    final latest = latestFolder(root, pc);
    final latestDate = await _backupDate(latest);
    if (latestDate != null) {
      result.add(
        BackupVersion(
          pc: pc,
          folder: latest,
          createdAt: latestDate,
          isLatest: true,
        ),
      );
    }
    final old = Directory(_join(pcFolder(root, pc).path, 'old'));
    if (await old.exists()) {
      await for (final entity in old.list(followLinks: false)) {
        if (entity is! Directory) continue;
        final date = await _backupDate(entity);
        if (date == null) continue;
        result.add(
          BackupVersion(
            pc: pc,
            folder: entity,
            createdAt: date,
            isLatest: false,
          ),
        );
      }
    }
    result.sort((left, right) {
      if (left.isLatest != right.isLatest) return left.isLatest ? -1 : 1;
      return right.createdAt.compareTo(left.createdAt);
    });
    return result;
  }

  Future<Map<String, DateTime>> fileModifiedDates(Directory folder) async {
    final dates = <String, DateTime>{};
    for (final name in configurationFileNames) {
      final file = File(_join(folder.path, name));
      if (await file.exists()) dates[name] = await file.lastModified();
    }
    return dates;
  }

  Future<DateTime?> _backupDate(Directory folder) async {
    final metadata = File(_join(folder.path, 'backup.json'));
    try {
      final decoded = jsonDecode(await metadata.readAsString());
      if (decoded is Map && decoded['created_at'] is String) {
        return DateTime.tryParse(decoded['created_at'] as String);
      }
    } on Object {
      // Older/manual backups can still be dated from their files.
    }
    DateTime? newest;
    for (final name in configurationFileNames) {
      final file = File(_join(folder.path, name));
      if (!await file.exists()) continue;
      final modified = await file.lastModified();
      if (newest == null || modified.isAfter(newest)) newest = modified;
    }
    return newest;
  }

  String _join(String parent, String child) =>
      '$parent${Platform.pathSeparator}$child';
}

class ConfigurationBackupService {
  ConfigurationBackupService({
    required this.projectRoot,
    String? historyRoot,
    DateTime Function()? clock,
    ConfigurationBackupRepository? repository,
    this.targetFolderResolver,
  }) : clock = clock ?? DateTime.now,
       historyRoot = historyRoot ?? '$projectRoot${Platform.pathSeparator}logs',
       repository = repository ?? const ConfigurationBackupRepository();

  final String projectRoot;
  final String historyRoot;
  final DateTime Function() clock;
  final ConfigurationBackupRepository repository;
  final String Function(String pc)? targetFolderResolver;

  Future<TargetConfigurationState> inspectTarget(String pc) async {
    final files = <String, bool>{};
    final modifiedAt = <String, DateTime>{};
    for (final name in configurationFileNames) {
      final file = File(_join(_targetFolder(pc), name));
      final exists = await file.exists();
      files[name] = exists;
      if (!exists) continue;
      final modified = await file.lastModified();
      modifiedAt[name] = modified;
    }
    return TargetConfigurationState(files: files, modifiedAt: modifiedAt);
  }

  Future<ConfigurationOperationResult> backup(
    String pc,
    String backupRoot, {
    required bool overwrite,
  }) async {
    final latest = repository.latestFolder(backupRoot, pc);
    final existing = await repository.hasBackup(backupRoot, pc);
    if (existing && !overwrite) {
      return ConfigurationOperationResult(
        pc: pc,
        success: true,
        copiedFiles: const [],
      );
    }

    final available = <String, File>{};
    for (final name in configurationFileNames) {
      final source = File(_join(_targetFolder(pc), name));
      if (await source.exists()) available[name] = source;
    }
    final hasCompleteAudio = audioConfigurationFileNames.every(
      available.containsKey,
    );
    final filesToCopy = <String>[
      if (hasCompleteAudio) ...audioConfigurationFileNames,
      if (available.containsKey(displayConfigurationFileName))
        displayConfigurationFileName,
    ];
    if (filesToCopy.isEmpty) {
      return ConfigurationOperationResult(
        pc: pc,
        success: false,
        copiedFiles: const [],
        error: 'The target has neither a complete audio configuration pair nor a display configuration.',
      );
    }

    try {
      if (existing) await _archiveLatest(backupRoot, pc, latest);
      await latest.create(recursive: true);
      final copied = <String>[];
      final modifiedAt = <String, String>{};
      for (final name in filesToCopy) {
        final source = available[name]!;
        final sourceModified = await source.lastModified();
        final destination = await source.copy(_join(latest.path, name));
        await destination.setLastModified(sourceModified);
        copied.add(name);
        modifiedAt[name] = sourceModified.toIso8601String();
      }
      final createdAt = clock();
      await File(_join(latest.path, 'backup.json')).writeAsString(
        const JsonEncoder.withIndent(' ').convert({
          'pc': pc,
          'created_at': createdAt.toIso8601String(),
          'files': copied,
          'file_modified_at': modifiedAt,
        }),
        flush: true,
      );
      await _writeBackupRecord(pc, createdAt, backupRoot, copied);
      return ConfigurationOperationResult(
        pc: pc,
        success: true,
        copiedFiles: copied,
      );
    } on Object catch (error) {
      return ConfigurationOperationResult(
        pc: pc,
        success: false,
        copiedFiles: const [],
        error: '$error',
      );
    }
  }

  Future<ConfigurationOperationResult> restore(
    String pc,
    BackupVersion version, {
    bool onlyMissing = false,
    Set<String>? includedFiles,
  }) async {
    final destinationFolder = Directory(_targetFolder(pc));
    try {
      await destinationFolder.create(recursive: true);
      final copied = <String>[];
      for (final name in configurationFileNames) {
        if (includedFiles != null && !includedFiles.contains(name)) continue;
        final source = File(_join(version.folder.path, name));
        if (!await source.exists()) continue;
        final destination = File(_join(destinationFolder.path, name));
        if (onlyMissing && await destination.exists()) continue;
        final temporary = File('${destination.path}.$pid.restore');
        try {
          await source.copy(temporary.path);
          if (await destination.exists()) await destination.delete();
          await temporary.rename(destination.path);
          copied.add(name);
        } finally {
          if (await temporary.exists()) await temporary.delete();
        }
      }
      if (copied.isEmpty) {
        return ConfigurationOperationResult(
          pc: pc,
          success: false,
          copiedFiles: const [],
          error: onlyMissing
              ? 'The target has no missing files that exist in this backup.'
              : 'The selected backup contains no configuration files.',
        );
      }
      return ConfigurationOperationResult(
        pc: pc,
        success: true,
        copiedFiles: copied,
      );
    } on Object catch (error) {
      return ConfigurationOperationResult(
        pc: pc,
        success: false,
        copiedFiles: const [],
        error: '$error',
      );
    }
  }

  Future<Map<String, bool>> filesIdentical(
    String pc,
    BackupVersion version,
  ) async {
    final results = <String, bool>{};
    for (final name in configurationFileNames) {
      final local = File(_join(_targetFolder(pc), name));
      final backup = File(_join(version.folder.path, name));
      if (!await local.exists() || !await backup.exists()) continue;
      results[name] = await _filesEqual(local, backup);
    }
    return results;
  }

  Future<bool> _filesEqual(File left, File right) async {
    if (await left.length() != await right.length()) return false;
    final leftBytes = await left.readAsBytes();
    final rightBytes = await right.readAsBytes();
    for (var index = 0; index < leftBytes.length; index++) {
      if (leftBytes[index] != rightBytes[index]) return false;
    }
    return true;
  }

  Future<void> _archiveLatest(String root, String pc, Directory latest) async {
    final versions = await repository.versions(root, pc);
    BackupVersion? current;
    for (final version in versions) {
      if (version.isLatest) {
        current = version;
        break;
      }
    }
    final date = current?.createdAt ?? clock();
    final oldRoot = Directory(_join(repository.pcFolder(root, pc).path, 'old'));
    await oldRoot.create(recursive: true);
    var archive = Directory(_join(oldRoot.path, _folderStamp(date)));
    var suffix = 2;
    while (await archive.exists()) {
      archive = Directory(_join(oldRoot.path, '${_folderStamp(date)}-$suffix'));
      suffix++;
    }
    await latest.rename(archive.path);
  }

  Future<void> _writeBackupRecord(
    String pc,
    DateTime createdAt,
    String backupRoot,
    List<String> files,
  ) async {
    final directory = Directory(_join(historyRoot, 'backup_records'));
    await directory.create(recursive: true);
    final file = File(
      _join(directory.path, '${repository.safePcName(pc)}.json'),
    );
    final safeHost = Platform.localHostname.replaceAll(
      RegExp(r'[^A-Za-z0-9_.-]+'),
      '_',
    );
    final temporary = File(
      '${file.path}.$safeHost.$pid.${DateTime.now().microsecondsSinceEpoch}.${Random.secure().nextInt(0x7fffffff)}.tmp',
    );
    try {
      await temporary.writeAsString(
        const JsonEncoder.withIndent(' ').convert({
          'pc': pc,
          'recorded_at': createdAt.toIso8601String(),
          'backup_root': backupRoot,
          'files': files,
        }),
        flush: true,
      );
      try {
        await temporary.rename(file.path);
      } on FileSystemException {
        await temporary.copy(file.path);
      }
    } finally {
      if (await temporary.exists()) await temporary.delete();
    }
  }

  String _targetFolder(String pc) {
    final customResolver = targetFolderResolver;
    if (customResolver != null) return customResolver(pc);
    final value = pc.trim();
    if (!RegExp(r'^[A-Za-z0-9_.-]+$').hasMatch(value)) {
      throw ArgumentError.value(pc, 'pc', 'contains unsupported characters');
    }
    final local = {
      'localhost',
      '.',
      '127.0.0.1',
      Platform.localHostname.toLowerCase(),
    }.contains(value.toLowerCase());
    return local
        ? r'C:\ProgramData\CTS'
        : r'\\' + value + r'\C$\ProgramData\CTS';
  }

  String _folderStamp(DateTime date) {
    final local = date.toLocal();
    return '${local.year.toString().padLeft(4, '0')}-'
        '${local.month.toString().padLeft(2, '0')}-'
        '${local.day.toString().padLeft(2, '0')}_'
        '${local.hour.toString().padLeft(2, '0')}-'
        '${local.minute.toString().padLeft(2, '0')}-'
        '${local.second.toString().padLeft(2, '0')}';
  }

  String _join(String parent, String child) {
    final normalized = child.replaceAll(
      RegExp(r'[\\/]'),
      Platform.pathSeparator,
    );
    return '$parent${Platform.pathSeparator}$normalized';
  }
}
