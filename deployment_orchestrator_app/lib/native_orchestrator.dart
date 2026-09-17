import 'dart:async';
import 'dart:convert';
import 'dart:io';

typedef LogSink = void Function(String line);
typedef ProgressSink = void Function(String target);

class CommandResult {
  const CommandResult({
    required this.exitCode,
    this.stdout = '',
    this.stderr = '',
  });

  final int exitCode;
  final String stdout;
  final String stderr;
}

abstract interface class CommandExecutor {
  bool get isCancelled;

  Future<CommandResult> run(
    String executable,
    List<String> arguments, {
    required String workingDirectory,
  });

  void cancel();
}

class ProcessCommandExecutor implements CommandExecutor {
  final Set<Process> _active = <Process>{};
  bool _cancelled = false;

  @override
  bool get isCancelled => _cancelled;

  @override
  Future<CommandResult> run(
    String executable,
    List<String> arguments, {
    required String workingDirectory,
  }) async {
    if (_cancelled) {
      return const CommandResult(exitCode: -1, stderr: 'Operation cancelled.');
    }
    final process = await Process.start(
      executable,
      arguments,
      workingDirectory: workingDirectory,
      runInShell: false,
    );
    _active.add(process);
    if (_cancelled) process.kill();
    try {
      final outputFuture = process.stdout
          .transform(const Utf8Decoder(allowMalformed: true))
          .join();
      final errorFuture = process.stderr
          .transform(const Utf8Decoder(allowMalformed: true))
          .join();
      final results = await Future.wait<Object>([
        process.exitCode,
        outputFuture,
        errorFuture,
      ]);
      return CommandResult(
        exitCode: results[0] as int,
        stdout: results[1] as String,
        stderr: results[2] as String,
      );
    } finally {
      _active.remove(process);
    }
  }

  @override
  void cancel() {
    _cancelled = true;
    for (final process in _active.toList()) {
      process.kill();
    }
  }
}

class DeploymentOptions {
  const DeploymentOptions({
    required this.audioRecall,
    required this.displayRecall,
    required this.bgInfoInstall,
    required this.desktopShortcuts,
    required this.bgInfoFolder,
  });

  final bool audioRecall;
  final bool displayRecall;
  final bool bgInfoInstall;
  final bool desktopShortcuts;
  final String bgInfoFolder;
}

class NativeOrchestrator {
  NativeOrchestrator({
    required this.projectRoot,
    required this.maxWorkers,
    CommandExecutor? executor,
    this.onLog,
    this.onMonitoringProgress,
    DateTime Function()? clock,
  }) : executor = executor ?? ProcessCommandExecutor(),
       _clock = clock ?? DateTime.now;

  final String projectRoot;
  final int maxWorkers;
  final CommandExecutor executor;
  final LogSink? onLog;
  final ProgressSink? onMonitoringProgress;
  final DateTime Function() _clock;

  IOSink? _logSink;

  void cancel() => executor.cancel();

  Future<Map<String, dynamic>> deploy(
    List<String> targets,
    DeploymentOptions options,
  ) async {
    _validate(targets, options: options);
    final normalizedTargets = _deduplicateTargets(targets);
    final logFile = await _openLog('');
    try {
      try {
        await _writeDeploymentRecords(normalizedTargets, options);
      } on FileSystemException catch (error) {
        _log('WARNING', 'Could not save deployment records: ${error.message}');
      }
      _log('INFO', 'Starting remote deployment run');
      _log(
        'INFO',
        'Targets: ${normalizedTargets.length}; maximum concurrent targets: $maxWorkers',
      );
      final results = await _parallelMap(normalizedTargets, (pc) async {
        try {
          return await _deployTarget(pc, options);
        } on Object catch (error, stackTrace) {
          _log('FATAL', '$pc: Unexpected deployment failure: $error');
          _log('DEBUG', '$pc: $stackTrace');
          return _fatalDeploymentResult(
            pc,
            'Unexpected deployment failure: $error',
          );
        }
      });
      _log('INFO', 'Remote deployment run complete');
      _log('INFO', 'Detailed log: ${logFile.absolute.path}');
      return {'log_file': logFile.absolute.path, 'pcs': results};
    } finally {
      await _closeLog();
    }
  }

  Future<Map<String, dynamic>> monitor(List<String> targets) async {
    _validate(targets);
    final normalizedTargets = _deduplicateTargets(targets);
    final logFile = await _openLog('monitor-');
    try {
      final results = await _parallelMap(normalizedTargets, (pc) async {
        try {
          return await _inspectTarget(pc);
        } on Object catch (error, stackTrace) {
          _log('ERROR', '$pc: Unexpected monitoring failure: $error');
          _log('DEBUG', '$pc: $stackTrace');
          final result = await _failedMonitoringResult(pc, '$error');
          onMonitoringProgress?.call(pc);
          return result;
        }
      });
      _log(
        'INFO',
        'Monitoring complete for ${normalizedTargets.length} target(s)',
      );
      return {'log_file': logFile.absolute.path, 'pcs': results};
    } finally {
      await _closeLog();
    }
  }

  Future<Map<String, dynamic>> uninstall(List<String> targets) async {
    _validate(targets);
    final normalizedTargets = _deduplicateTargets(targets);
    final logFile = await _openLog('uninstall-');
    try {
      _log('INFO', 'Starting remote uninstall run');
      _log(
        'INFO',
        'Targets: ${normalizedTargets.length}; maximum concurrent targets: $maxWorkers',
      );
      final results = await _parallelMap(normalizedTargets, (pc) async {
        try {
          return await _uninstallTarget(pc);
        } on Object catch (error, stackTrace) {
          _log('ERROR', '$pc: Unexpected uninstall failure: $error');
          _log('DEBUG', '$pc: $stackTrace');
          return <String, dynamic>{
            'pc': pc,
            'success': false,
            'error': 'Unexpected uninstall failure: $error',
          };
        }
      });
      _log('INFO', 'Remote uninstall run complete');
      _log('INFO', 'Detailed log: ${logFile.absolute.path}');
      return {'log_file': logFile.absolute.path, 'pcs': results};
    } finally {
      await _closeLog();
    }
  }

  void _validate(List<String> targets, {DeploymentOptions? options}) {
    if (maxWorkers < 1) {
      throw ArgumentError.value(maxWorkers, 'maxWorkers', 'must be at least 1');
    }
    if (_deduplicateTargets(targets).isEmpty) {
      throw ArgumentError.value(targets, 'targets', 'must not be empty');
    }
    if (options?.bgInfoInstall == true &&
        options!.bgInfoFolder.trim().isEmpty) {
      throw ArgumentError(
        'BGInfo folder cannot be empty when BGInfo is enabled.',
      );
    }
  }

  List<String> _deduplicateTargets(List<String> targets) {
    return targets
        .map((target) => target.trim())
        .where((target) => target.isNotEmpty && !target.startsWith('#'))
        .toSet()
        .toList();
  }

  Future<File> _openLog(String prefix) async {
    final logs = Directory(_join(projectRoot, 'logs'));
    await logs.create(recursive: true);
    final now = _clock();
    final stamp =
        '${now.year.toString().padLeft(4, '0')}-'
        '${now.month.toString().padLeft(2, '0')}-'
        '${now.day.toString().padLeft(2, '0')}_'
        '${now.hour.toString().padLeft(2, '0')}-'
        '${now.minute.toString().padLeft(2, '0')}-'
        '${now.second.toString().padLeft(2, '0')}-'
        '${now.millisecond.toString().padLeft(3, '0')}';
    final file = File(_join(logs.path, '$prefix$stamp.log'));
    _logSink = file.openWrite(mode: FileMode.writeOnly);
    return file;
  }

  Future<void> _closeLog() async {
    final sink = _logSink;
    _logSink = null;
    await sink?.flush();
    await sink?.close();
  }

  void _log(String level, String message) {
    final now = _clock().toIso8601String();
    final line = '$now $level $message';
    _logSink?.writeln(line);
    onLog?.call(line);
  }

  Future<List<T>> _parallelMap<S, T>(
    List<S> items,
    Future<T> Function(S item) operation,
  ) async {
    final results = List<T?>.filled(items.length, null);
    var next = 0;

    Future<void> worker() async {
      while (true) {
        final index = next++;
        if (index >= items.length) return;
        results[index] = await operation(items[index]);
      }
    }

    final workerCount = maxWorkers < items.length ? maxWorkers : items.length;
    await Future.wait(List.generate(workerCount, (_) => worker()));
    return results.cast<T>();
  }

  Future<Map<String, dynamic>> _deployTarget(
    String pc,
    DeploymentOptions options,
  ) async {
    final result = <String, dynamic>{
      'pc': pc,
      'highest_severity': 'info',
      'issues': <Map<String, String>>[],
      'scripts': <Map<String, dynamic>>[],
    };

    if (executor.isCancelled) {
      _recordIssue(result, 'fatal', 'Deployment cancelled', 'Orchestrator');
      return result;
    }

    _log('INFO', '$pc: Queuing configuration check');
    final ping = await _safeRun(
      'powershell.exe',
      ['ping', '-n', '1', pc],
      pc: pc,
      action: 'Ping test',
    );
    if (ping == null || ping.exitCode != 0) {
      _log('WARNING', '$pc: Ping test failed');
      _recordIssue(result, 'warning', 'Ping test failed', 'Connectivity');
      _log('INFO', '$pc: Deployment complete with connectivity issues');
      return result;
    }

    final winRm = await _safeRun(
      'powershell.exe',
      ['Invoke-Command', '-ComputerName', pc, '-ScriptBlock', '{1}'],
      pc: pc,
      action: 'WinRM test',
    );
    if (winRm == null || winRm.exitCode != 0) {
      _log('ERROR', '$pc: WinRM test failed');
      _recordIssue(result, 'error', 'WinRM test failed', 'Connectivity');
      _log('INFO', '$pc: Deployment complete with connectivity issues');
      return result;
    }

    final scripts = <(String, bool)>[
      ('installer_scripts\\InstallAudioDeviceCmdlets.ps1', options.audioRecall),
      ('installer_scripts\\InstallDisplayConfig.ps1', options.displayRecall),
      ('installer_scripts\\InstallBGInfo.ps1', options.bgInfoInstall),
      ('installer_scripts\\Cleanup.ps1', true),
      ('installer_scripts\\AddShortcuts.ps1', options.desktopShortcuts),
    ];

    for (final (script, enabled) in scripts) {
      if (!enabled || executor.isCancelled) continue;
      final scriptResult = <String, dynamic>{
        'name': script,
        'severity': 'info',
        'messages': <String>[],
      };
      (result['scripts'] as List<Map<String, dynamic>>).add(scriptResult);
      _log('INFO', '$pc: Running $script');
      final arguments = <String>[
        '-File',
        _join(projectRoot, script),
        pc,
        'false',
        if (script.toLowerCase().contains('bginfo')) options.bgInfoFolder,
      ];
      final completed = await _safeRun(
        'powershell.exe',
        arguments,
        pc: pc,
        action: script,
      );
      if (completed == null) {
        _recordScriptEvent(
          result,
          scriptResult,
          'fatal',
          'Could not start script',
          script,
        );
        continue;
      }
      _recordProcessOutput(result, scriptResult, pc, script, completed);
      if (completed.exitCode == 0) {
        _log('INFO', '$pc: Completed $script');
      } else {
        final message = 'Failed with exit code ${completed.exitCode}';
        _recordScriptEvent(result, scriptResult, 'error', message, script);
        _log('ERROR', '$pc: $script ${message.toLowerCase()}');
      }
    }
    _log('INFO', '$pc: Deployment complete');
    return result;
  }

  Future<Map<String, dynamic>> _uninstallTarget(String pc) async {
    if (executor.isCancelled) {
      return <String, dynamic>{
        'pc': pc,
        'success': false,
        'error': 'Uninstall cancelled.',
      };
    }

    _log('INFO', '$pc: Uninstall started');
    final completed = await _safeRun(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        _join(projectRoot, 'utility_scripts\\uninstall.ps1'),
        pc,
        'true',
      ],
      pc: pc,
      action: 'Uninstall',
    );
    if (completed == null) {
      return <String, dynamic>{
        'pc': pc,
        'success': false,
        'error': 'Could not start uninstall script.',
      };
    }
    for (final (output, defaultLevel) in [
      (completed.stdout, 'INFO'),
      (completed.stderr, 'ERROR'),
    ]) {
      for (final line in const LineSplitter().convert(output)) {
        if (line.trim().isEmpty) continue;
        _log(defaultLevel, '$pc: ${line.trim()}');
      }
    }
    if (completed.exitCode != 0) {
      final error = completed.stderr.trim().isNotEmpty
          ? completed.stderr.trim()
          : 'Uninstall script exited with code ${completed.exitCode}.';
      _log('ERROR', '$pc: Uninstall failed');
      return <String, dynamic>{'pc': pc, 'success': false, 'error': error};
    }

    try {
      await _writeUninstallRecord(pc);
    } on FileSystemException catch (error) {
      _log('WARNING', '$pc: Could not save uninstall record: ${error.message}');
    }
    _log('INFO', '$pc: Uninstall complete');
    return <String, dynamic>{'pc': pc, 'success': true, 'error': ''};
  }

  Map<String, dynamic> _fatalDeploymentResult(String pc, String message) {
    return <String, dynamic>{
      'pc': pc,
      'highest_severity': 'fatal',
      'issues': <Map<String, String>>[
        {'component': 'Orchestrator', 'severity': 'fatal', 'message': message},
      ],
      'scripts': <Map<String, dynamic>>[],
    };
  }

  Future<CommandResult?> _safeRun(
    String executable,
    List<String> arguments, {
    required String pc,
    required String action,
  }) async {
    try {
      return await executor.run(
        executable,
        arguments,
        workingDirectory: projectRoot,
      );
    } on ProcessException catch (error) {
      _log('FATAL', '$pc: $action: Could not start process: ${error.message}');
      return null;
    } on OSError catch (error) {
      _log('FATAL', '$pc: $action: Could not start process: $error');
      return null;
    }
  }

  void _recordProcessOutput(
    Map<String, dynamic> pcResult,
    Map<String, dynamic> scriptResult,
    String pc,
    String script,
    CommandResult completed,
  ) {
    for (final (output, defaultLevel) in [
      (completed.stdout, 'INFO'),
      (completed.stderr, 'ERROR'),
    ]) {
      for (final rawLine in const LineSplitter().convert(output)) {
        if (rawLine.isEmpty) continue;
        final parsed = RegExp(
          r'^\s*(DEBUG|INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\s*[:-]?\s*(.*)$',
          caseSensitive: false,
        ).firstMatch(rawLine);
        final level = (parsed?.group(1) ?? defaultLevel).toUpperCase();
        final message = parsed?.group(2)?.isNotEmpty == true
            ? parsed!.group(2)!
            : rawLine;
        final severity = _normalizeSeverity(level);
        if (severity != 'info') {
          _recordScriptEvent(
            pcResult,
            scriptResult,
            severity,
            message,
            script,
            writeLog: false,
          );
        }
        _log(level == 'WARN' ? 'WARNING' : level, '$pc: $script: $message');
      }
    }
  }

  void _recordIssue(
    Map<String, dynamic> result,
    String severity,
    String message,
    String component,
  ) {
    _raiseSeverity(result, severity);
    (result['issues'] as List<Map<String, String>>).add({
      'component': component,
      'severity': severity,
      'message': message,
    });
  }

  void _recordScriptEvent(
    Map<String, dynamic> result,
    Map<String, dynamic> scriptResult,
    String severity,
    String message,
    String script, {
    bool writeLog = true,
  }) {
    _raiseSeverity(result, severity);
    if (_severityRank(severity) >
        _severityRank(scriptResult['severity'] as String)) {
      scriptResult['severity'] = severity;
    }
    (scriptResult['messages'] as List<String>).add(message);
    if (writeLog) {
      _log(severity.toUpperCase(), '${result['pc']}: $script: $message');
    }
  }

  void _raiseSeverity(Map<String, dynamic> result, String severity) {
    if (_severityRank(severity) >
        _severityRank(result['highest_severity'] as String)) {
      result['highest_severity'] = severity;
    }
  }

  int _severityRank(String severity) => switch (severity.toLowerCase()) {
    'fatal' => 3,
    'error' => 2,
    'warning' => 1,
    _ => 0,
  };

  String _normalizeSeverity(String severity) =>
      switch (severity.toUpperCase()) {
        'WARN' || 'WARNING' => 'warning',
        'ERROR' => 'error',
        'FATAL' || 'CRITICAL' => 'fatal',
        _ => 'info',
      };

  Future<void> _writeDeploymentRecords(
    List<String> targets,
    DeploymentOptions options,
  ) async {
    final directory = Directory(
      _join(_join(projectRoot, 'logs'), 'deployment_records'),
    );
    await directory.create(recursive: true);
    final recordedAt = _clock().toIso8601String();
    for (final pc in targets) {
      final file = File(_join(directory.path, _recordName(pc)));
      final temporary = File('${file.path}.$pid.tmp');
      final payload = {
        'pc': pc,
        'recorded_at': recordedAt,
        'audio_recall': options.audioRecall,
        'display_recall': options.displayRecall,
        'bginfo_install': options.bgInfoInstall,
        'desktop_shortcuts': options.desktopShortcuts,
      };
      await temporary.writeAsString(
        const JsonEncoder.withIndent('  ').convert(payload),
      );
      await temporary.rename(file.path);
      final uninstallRecord = File(
        _join(
          _join(_join(projectRoot, 'logs'), 'uninstall_records'),
          _recordName(pc),
        ),
      );
      if (await uninstallRecord.exists()) await uninstallRecord.delete();
    }
  }

  Future<void> _writeUninstallRecord(String pc) async {
    final directory = Directory(
      _join(_join(projectRoot, 'logs'), 'uninstall_records'),
    );
    await directory.create(recursive: true);
    final file = File(_join(directory.path, _recordName(pc)));
    final temporary = File('${file.path}.$pid.tmp');
    await temporary.writeAsString(
      const JsonEncoder.withIndent(' ')
          .convert({'pc': pc, 'recorded_at': _clock().toIso8601String()}),
    );
    await temporary.rename(file.path);
  }

  Future<Map<String, dynamic>> _inspectTarget(String pc) async {
    _log('INFO', '$pc: Monitoring started');
    Map<String, dynamic> result;
    if (executor.isCancelled) {
      result = await _failedMonitoringResult(pc, 'Monitoring cancelled.');
    } else {
      final completed = await _safeRun(
        'powershell.exe',
        [
          '-NoProfile',
          '-NonInteractive',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          _join(projectRoot, 'utility_scripts\\MonitorTarget.ps1'),
          '-ComputerName',
          pc,
        ],
        pc: pc,
        action: 'Monitoring',
      );
      final outputLines =
          completed?.stdout
              .split(RegExp(r'[\r\n]+'))
              .map((line) => line.trim())
              .where((line) => line.isNotEmpty)
              .toList() ??
          const <String>[];
      if (completed == null || completed.exitCode != 0 || outputLines.isEmpty) {
        final message = completed?.stderr.trim().isNotEmpty == true
            ? completed!.stderr.trim()
            : 'PowerShell returned no monitoring result.';
        _log('ERROR', '$pc: Monitoring failed: $message');
        result = await _failedMonitoringResult(pc, message);
      } else {
        try {
          result = Map<String, dynamic>.from(
            jsonDecode(outputLines.last) as Map,
          );
          result['deployment_intent'] = await _loadDeploymentIntent(pc);
          result['uninstall_recorded_at'] = await _loadUninstallRecord(pc);
          final status = result['online'] != true
              ? 'offline'
              : result['winrm'] == true
              ? 'complete'
              : 'error';
          _log('INFO', '$pc: Monitoring complete ($status)');
        } on FormatException catch (error) {
          final message = 'Invalid PowerShell monitoring result: $error';
          _log('ERROR', '$pc: $message');
          result = await _failedMonitoringResult(pc, message);
        } on TypeError catch (error) {
          final message = 'Invalid PowerShell monitoring result: $error';
          _log('ERROR', '$pc: $message');
          result = await _failedMonitoringResult(pc, message);
        }
      }
    }
    onMonitoringProgress?.call(pc);
    return result;
  }

  Future<Map<String, dynamic>> _failedMonitoringResult(
    String pc,
    String message,
  ) async {
    return {
      'pc': pc,
      'online': false,
      'winrm': false,
      'error': message,
      'cts_deployed': false,
      'audio_configured': false,
      'display_configured': false,
      'logout_shortcut': false,
      'reboot_shortcut': false,
      'bginfo_deployed': false,
      'bginfo_executable': false,
      'bginfo_profile': false,
      'bginfo_background': false,
      'bginfo_startup': false,
      'bginfo_startup_method': '',
      'audio_device_cmdlets_versions': <String>[],
      'display_config_versions': <String>[],
      'deployment_intent': await _loadDeploymentIntent(pc),
      'uninstall_recorded_at': await _loadUninstallRecord(pc),
    };
  }

  Future<Map<String, dynamic>?> _loadDeploymentIntent(String pc) async {
    final file = File(
      _join(
        _join(_join(projectRoot, 'logs'), 'deployment_records'),
        _recordName(pc),
      ),
    );
    try {
      final decoded = jsonDecode(await file.readAsString());
      return decoded is Map<String, dynamic> ? decoded : null;
    } on FileSystemException {
      return null;
    } on FormatException {
      return null;
    }
  }

  Future<String?> _loadUninstallRecord(String pc) async {
    final file = File(
      _join(
        _join(_join(projectRoot, 'logs'), 'uninstall_records'),
        _recordName(pc),
      ),
    );
    try {
      final decoded = jsonDecode(await file.readAsString());
      return decoded is Map<String, dynamic>
          ? decoded['recorded_at'] as String?
          : null;
    } on FileSystemException {
      return null;
    } on FormatException {
      return null;
    }
  }

  String _recordName(String pc) {
    final safe = pc
        .replaceAll(RegExp(r'[^A-Za-z0-9_.-]+'), '_')
        .replaceAll(RegExp(r'^[._]+|[._]+$'), '');
    return '${safe.isEmpty ? 'unknown' : safe.toLowerCase()}.json';
  }

  String _join(String parent, String child) {
    final normalizedChild = child.replaceAll(
      RegExp(r'[\\/]'),
      Platform.pathSeparator,
    );
    return '$parent${Platform.pathSeparator}$normalizedChild';
  }
}
