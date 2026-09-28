import 'dart:io';

import 'package:deployment_orchestrator_app/configuration_backup.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  late Directory workspace;
  late Directory target;
  late Directory backups;

  setUp(() async {
    workspace = await Directory.systemTemp.createTemp('config-backup-test-');
    target = Directory('${workspace.path}${Platform.pathSeparator}target');
    backups = Directory('${workspace.path}${Platform.pathSeparator}backups');
    await target.create(recursive: true);
  });

  tearDown(() async {
    if (await workspace.exists()) await workspace.delete(recursive: true);
  });

  Future<void> writeTarget(String value) async {
    for (final name in configurationFileNames) {
      await File('${target.path}${Platform.pathSeparator}$name')
          .writeAsString('$name-$value');
    }
  }

  test(
    'backs up all configuration files and records the latest version',
    () async {
      await writeTarget('one');
      final createdAt = DateTime.utc(2026, 9, 23, 12, 30);
      final sharedHistory = Directory(
        '${workspace.path}${Platform.pathSeparator}shared-history',
      );
      final service = ConfigurationBackupService(
        projectRoot: workspace.path,
        historyRoot: sharedHistory.path,
        clock: () => createdAt,
        targetFolderResolver: (_) => target.path,
      );

      final result = await service.backup(
        'PC-001',
        backups.path,
        overwrite: false,
      );

      expect(result.success, isTrue);
      expect(result.copiedFiles, unorderedEquals(configurationFileNames));
      final versions = await const ConfigurationBackupRepository().versions(
        backups.path,
        'PC-001',
      );
      expect(versions, hasLength(1));
      expect(versions.single.isLatest, isTrue);
      expect(versions.single.createdAt, createdAt);
      expect(
        File(
          '${sharedHistory.path}${Platform.pathSeparator}'
          'backup_records${Platform.pathSeparator}pc-001.json',
        ).existsSync(),
        isTrue,
      );
    },
  );

  test('moves an overwritten latest backup into a dated old folder', () async {
    var now = DateTime(2026, 9, 22, 8);
    final service = ConfigurationBackupService(
      projectRoot: workspace.path,
      clock: () => now,
      targetFolderResolver: (_) => target.path,
    );
    await writeTarget('old');
    expect(
      (await service.backup('PC-002', backups.path, overwrite: false)).success,
      isTrue,
    );
    now = DateTime(2026, 9, 23, 9);
    await writeTarget('new');

    final result = await service.backup(
      'PC-002',
      backups.path,
      overwrite: true,
    );

    expect(result.success, isTrue);
    final versions = await const ConfigurationBackupRepository().versions(
      backups.path,
      'PC-002',
    );
    expect(versions, hasLength(2));
    expect(versions.first.isLatest, isTrue);
    expect(versions.last.isLatest, isFalse);
    expect(versions.last.folder.path, contains('2026-09-22_08-00-00'));
    expect(
      await File(
        '${versions.first.folder.path}${Platform.pathSeparator}audio_levels.json',
      ).readAsString(),
      'audio_levels.json-new',
    );
  });

  test(
    'restore missing leaves existing target configuration untouched',
    () async {
      await writeTarget('backup');
      final service = ConfigurationBackupService(
        projectRoot: workspace.path,
        clock: () => DateTime.utc(2026, 9, 23),
        targetFolderResolver: (_) => target.path,
      );
      await service.backup('PC-003', backups.path, overwrite: false);
      for (final name in configurationFileNames) {
        await File('${target.path}${Platform.pathSeparator}$name').delete();
      }
      await File('${target.path}${Platform.pathSeparator}audio_levels.json')
          .writeAsString('keep-local');
      final version = (await const ConfigurationBackupRepository().versions(
        backups.path,
        'PC-003',
      )).first;

      final result = await service.restore(
        'PC-003',
        version,
        onlyMissing: true,
      );

      expect(result.success, isTrue);
      expect(result.copiedFiles, isNot(contains('audio_levels.json')));
      expect(
        await File('${target.path}${Platform.pathSeparator}audio_levels.json')
            .readAsString(),
        'keep-local',
      );
      expect(
        File(
          '${target.path}${Platform.pathSeparator}display_config_profile.xml',
        ).existsSync(),
        isTrue,
      );
    },
  );

  test('backs up display but skips an incomplete audio pair', () async {
    await writeTarget('complete');
    final service = ConfigurationBackupService(
      projectRoot: workspace.path,
      targetFolderResolver: (_) => target.path,
    );
    await service.backup('PC-004', backups.path, overwrite: false);
    await File('${target.path}${Platform.pathSeparator}audio_device_list.json')
        .delete();

    final result = await service.backup(
      'PC-004',
      backups.path,
      overwrite: true,
    );

    expect(result.success, isTrue);
    expect(result.copiedFiles, [displayConfigurationFileName]);
    final versions = await const ConfigurationBackupRepository().versions(
      backups.path,
      'PC-004',
    );
    expect(versions, hasLength(2));
    expect(
      File(
        '${versions.first.folder.path}${Platform.pathSeparator}audio_device_list.json',
      ).existsSync(),
      isFalse,
    );
    expect(
      File(
        '${versions.first.folder.path}${Platform.pathSeparator}$displayConfigurationFileName',
      ).existsSync(),
      isTrue,
    );
  });

  test('backs up the audio pair when the display profile is missing', () async {
    await writeTarget('audio-only');
    await File(
      '${target.path}${Platform.pathSeparator}$displayConfigurationFileName',
    ).delete();
    final service = ConfigurationBackupService(
      projectRoot: workspace.path,
      targetFolderResolver: (_) => target.path,
    );

    final result = await service.backup(
      'PC-005',
      backups.path,
      overwrite: false,
    );

    expect(result.success, isTrue);
    expect(result.copiedFiles, unorderedEquals(audioConfigurationFileNames));
  });

  test(
    'fails only when neither complete configuration group is available',
    () async {
      await writeTarget('incomplete');
      await File(
        '${target.path}${Platform.pathSeparator}audio_device_list.json',
      ).delete();
      await File(
        '${target.path}${Platform.pathSeparator}$displayConfigurationFileName',
      ).delete();
      final service = ConfigurationBackupService(
        projectRoot: workspace.path,
        targetFolderResolver: (_) => target.path,
      );

      final result = await service.backup(
        'PC-006',
        backups.path,
        overwrite: false,
      );

      expect(result.success, isFalse);
      expect(result.error, contains('neither a complete audio configuration'));
    },
  );

  test('restores every available file from a partial backup', () async {
    await writeTarget('partial');
    await File(
      '${target.path}${Platform.pathSeparator}$displayConfigurationFileName',
    ).delete();
    final service = ConfigurationBackupService(
      projectRoot: workspace.path,
      targetFolderResolver: (_) => target.path,
    );
    await service.backup('PC-007', backups.path, overwrite: false);
    for (final name in audioConfigurationFileNames) {
      await File('${target.path}${Platform.pathSeparator}$name').delete();
    }
    final version = (await const ConfigurationBackupRepository().versions(
      backups.path,
      'PC-007',
    )).first;

    final result = await service.restore('PC-007', version);

    expect(result.success, isTrue);
    expect(result.copiedFiles, unorderedEquals(audioConfigurationFileNames));
    expect(
      File(
        '${target.path}${Platform.pathSeparator}$displayConfigurationFileName',
      ).existsSync(),
      isFalse,
    );
  });

  test('checks file contents when comparing a target with a backup', () async {
    await writeTarget('same');
    final service = ConfigurationBackupService(
      projectRoot: workspace.path,
      targetFolderResolver: (_) => target.path,
    );
    await service.backup('PC-008', backups.path, overwrite: false);
    final version = (await const ConfigurationBackupRepository().versions(
      backups.path,
      'PC-008',
    )).first;

    var identical = await service.filesIdentical('PC-008', version);
    expect(identical.values, everyElement(isTrue));

    final changed = File(
      '${target.path}${Platform.pathSeparator}audio_levels.json',
    );
    final originalModified = await changed.lastModified();
    await changed.writeAsString('different');
    await changed.setLastModified(originalModified);
    identical = await service.filesIdentical('PC-008', version);
    expect(identical['audio_levels.json'], isFalse);
  });
}
