import 'dart:async';
import 'dart:convert';
import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

void main() {
  runApp(const DeploymentOrchestratorApp());
}

class DeploymentOrchestratorApp extends StatefulWidget {
  const DeploymentOrchestratorApp({super.key});

  @override
  State<DeploymentOrchestratorApp> createState() =>
      _DeploymentOrchestratorAppState();
}

class _DeploymentOrchestratorAppState extends State<DeploymentOrchestratorApp> {
  bool _darkMode = true;

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Deployment Orchestrator',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xff315c8c)),
        useMaterial3: true,
      ),
      darkTheme: ThemeData(
        brightness: Brightness.dark,
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xff78a9dc),
          brightness: Brightness.dark,
        ),
        useMaterial3: true,
      ),
      themeMode: _darkMode ? ThemeMode.dark : ThemeMode.light,
      home: DeploymentPage(
        darkMode: _darkMode,
        onToggleTheme: () => setState(() => _darkMode = !_darkMode),
      ),
    );
  }
}

enum TargetSource { file, direct }

enum PcProgress { queued, running, complete, offline, warning, error, fatal }

enum MonitoringFilter {
  all,
  failedChecks,
  offline,
  winRmUnavailable,
  missingDeployment,
  missingAudioConfiguration,
  missingDisplayConfiguration,
  missingLogoutShortcut,
  missingRebootShortcut,
  missingBgInfo,
  missingAudioDeviceCmdlets,
  missingDisplayConfig,
}

enum MonitoringComponentStatus { present, missing, notDeployed, unknown }

class DeploymentIntent {
  const DeploymentIntent({
    required this.recordedAt,
    required this.audioRecall,
    required this.displayRecall,
    required this.bgInfoInstall,
    required this.desktopShortcuts,
  });

  factory DeploymentIntent.fromJson(Map<String, dynamic> json) {
    return DeploymentIntent(
      recordedAt: json['recorded_at'] as String? ?? '',
      audioRecall: json['audio_recall'] as bool? ?? true,
      displayRecall: json['display_recall'] as bool? ?? true,
      bgInfoInstall: json['bginfo_install'] as bool? ?? true,
      desktopShortcuts: json['desktop_shortcuts'] as bool? ?? true,
    );
  }

  final String recordedAt;
  final bool audioRecall;
  final bool displayRecall;
  final bool bgInfoInstall;
  final bool desktopShortcuts;
}

class MonitoringReport {
  const MonitoringReport({required this.logFile, required this.pcs});

  factory MonitoringReport.fromJson(Map<String, dynamic> json) {
    return MonitoringReport(
      logFile: json['log_file'] as String? ?? '',
      pcs: (json['pcs'] as List<dynamic>? ?? const [])
          .map(
            (item) => MonitoringPcResult.fromJson(item as Map<String, dynamic>),
          )
          .toList(),
    );
  }

  final String logFile;
  final List<MonitoringPcResult> pcs;
}

class MonitoringPcResult {
  const MonitoringPcResult({
    required this.pc,
    required this.online,
    required this.winRm,
    required this.error,
    required this.ctsDeployed,
    required this.audioConfigured,
    required this.displayConfigured,
    required this.logoutShortcut,
    required this.rebootShortcut,
    required this.bgInfoDeployed,
    required this.bgInfoExecutable,
    required this.bgInfoProfile,
    required this.bgInfoBackground,
    required this.bgInfoStartup,
    required this.bgInfoStartupMethod,
    required this.audioDeviceCmdletsVersions,
    required this.displayConfigVersions,
    required this.deploymentIntent,
  });

  factory MonitoringPcResult.fromJson(Map<String, dynamic> json) {
    List<String> versions(String key) =>
        (json[key] as List<dynamic>? ?? const [])
            .map((item) => '$item')
            .toList();

    return MonitoringPcResult(
      pc: json['pc'] as String? ?? 'Unknown PC',
      online: json['online'] as bool? ?? false,
      winRm: json['winrm'] as bool? ?? false,
      error: json['error'] as String? ?? '',
      ctsDeployed: json['cts_deployed'] as bool? ?? false,
      audioConfigured: json['audio_configured'] as bool? ?? false,
      displayConfigured: json['display_configured'] as bool? ?? false,
      logoutShortcut: json['logout_shortcut'] as bool? ?? false,
      rebootShortcut: json['reboot_shortcut'] as bool? ?? false,
      bgInfoDeployed: json['bginfo_deployed'] as bool? ?? false,
      bgInfoExecutable: json['bginfo_executable'] as bool? ?? false,
      bgInfoProfile: json['bginfo_profile'] as bool? ?? false,
      bgInfoBackground: json['bginfo_background'] as bool? ?? false,
      bgInfoStartup: json['bginfo_startup'] as bool? ?? false,
      bgInfoStartupMethod: json['bginfo_startup_method'] as String? ?? '',
      audioDeviceCmdletsVersions: versions('audio_device_cmdlets_versions'),
      displayConfigVersions: versions('display_config_versions'),
      deploymentIntent: json['deployment_intent'] is Map<String, dynamic>
          ? DeploymentIntent.fromJson(
              json['deployment_intent'] as Map<String, dynamic>,
            )
          : null,
    );
  }

  final String pc;
  final bool online;
  final bool winRm;
  final String error;
  final bool ctsDeployed;
  final bool audioConfigured;
  final bool displayConfigured;
  final bool logoutShortcut;
  final bool rebootShortcut;
  final bool bgInfoDeployed;
  final bool bgInfoExecutable;
  final bool bgInfoProfile;
  final bool bgInfoBackground;
  final bool bgInfoStartup;
  final String bgInfoStartupMethod;
  final List<String> audioDeviceCmdletsVersions;
  final List<String> displayConfigVersions;
  final DeploymentIntent? deploymentIntent;
}

List<TextSpan> buildSeveritySpans(String text) {
  final severityPattern = RegExp(
    r'\b(INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\b',
    caseSensitive: false,
  );
  final spans = <TextSpan>[];
  var cursor = 0;

  for (final match in severityPattern.allMatches(text)) {
    if (match.start > cursor) {
      spans.add(TextSpan(text: text.substring(cursor, match.start)));
    }
    final severity = match.group(0)!;
    final normalized = severity.toUpperCase();
    final color = switch (normalized) {
      'INFO' => const Color(0xff22c55e),
      'WARNING' || 'WARN' => const Color(0xffff9800),
      'ERROR' || 'FATAL' || 'CRITICAL' => const Color(0xffef4444),
      _ => null,
    };
    spans.add(
      TextSpan(
        text: severity,
        style: TextStyle(
          color: color,
          fontWeight: normalized == 'FATAL' || normalized == 'CRITICAL'
              ? FontWeight.bold
              : null,
        ),
      ),
    );
    cursor = match.end;
  }

  if (cursor < text.length) {
    spans.add(TextSpan(text: text.substring(cursor)));
  }
  return spans;
}

String filterDetailedLog(String output, {String? selectedPc}) {
  final pcMarker = selectedPc == null
      ? null
      : RegExp('(?:^|\\s)${RegExp.escape(selectedPc)}:');
  return output
      .split('\n')
      .where((line) => pcMarker == null || pcMarker.hasMatch(line))
      .join('\n');
}

List<String> filterPcChoices(Iterable<String> pcs, String query) {
  final normalizedQuery = query.trim().toLowerCase();
  return pcs
      .where(
        (pc) =>
            normalizedQuery.isEmpty ||
            pc.toLowerCase().contains(normalizedQuery),
      )
      .toList();
}

const allDeploymentScriptNames = <String>[
  'InstallAudioDeviceCmdlets.ps1',
  'InstallDisplayConfig.ps1',
  'InstallBGInfo.ps1',
  'Cleanup.ps1',
  'AddShortcuts.ps1',
];

String scriptFileName(String path) => path.split(RegExp(r'[\\/]')).last;

List<ScriptDeploymentResult> completeScriptStatusList(
  Iterable<ScriptDeploymentResult> reportedScripts,
) {
  final reportedByName = <String, ScriptDeploymentResult>{
    for (final script in reportedScripts)
      scriptFileName(script.name).toLowerCase(): script,
  };
  return [
    for (final scriptName in allDeploymentScriptNames)
      reportedByName[scriptName.toLowerCase()] ??
          ScriptDeploymentResult(
            name: scriptName,
            severity: 'not_deployed',
            messages: const [],
          ),
  ];
}

String truncateComputerName(String name) {
  if (name.length <= 15) return name;
  return '${name.substring(0, 12)}...';
}

class DeploymentReport {
  const DeploymentReport({required this.logFile, required this.pcs});

  factory DeploymentReport.fromJson(Map<String, dynamic> json) {
    return DeploymentReport(
      logFile: json['log_file'] as String? ?? '',
      pcs: (json['pcs'] as List<dynamic>? ?? const [])
          .map(
            (item) => PcDeploymentResult.fromJson(item as Map<String, dynamic>),
          )
          .toList(),
    );
  }

  final String logFile;
  final List<PcDeploymentResult> pcs;
}

class PcDeploymentResult {
  const PcDeploymentResult({
    required this.pc,
    required this.highestSeverity,
    required this.issues,
    required this.scripts,
  });

  factory PcDeploymentResult.fromJson(Map<String, dynamic> json) {
    return PcDeploymentResult(
      pc: json['pc'] as String? ?? 'Unknown PC',
      highestSeverity: json['highest_severity'] as String? ?? 'fatal',
      issues: (json['issues'] as List<dynamic>? ?? const [])
          .map((item) => DeploymentIssue.fromJson(item as Map<String, dynamic>))
          .toList(),
      scripts: (json['scripts'] as List<dynamic>? ?? const [])
          .map(
            (item) =>
                ScriptDeploymentResult.fromJson(item as Map<String, dynamic>),
          )
          .toList(),
    );
  }

  final String pc;
  final String highestSeverity;
  final List<DeploymentIssue> issues;
  final List<ScriptDeploymentResult> scripts;

  bool get hasProblems => highestSeverity.toLowerCase() != 'info';
}

class DeploymentIssue {
  const DeploymentIssue({
    required this.component,
    required this.severity,
    required this.message,
  });

  factory DeploymentIssue.fromJson(Map<String, dynamic> json) {
    return DeploymentIssue(
      component: json['component'] as String? ?? 'Deployment',
      severity: json['severity'] as String? ?? 'error',
      message: json['message'] as String? ?? 'Unknown failure',
    );
  }

  final String component;
  final String severity;
  final String message;
}

class ScriptDeploymentResult {
  const ScriptDeploymentResult({
    required this.name,
    required this.severity,
    required this.messages,
  });

  factory ScriptDeploymentResult.fromJson(Map<String, dynamic> json) {
    return ScriptDeploymentResult(
      name: json['name'] as String? ?? 'Unknown script',
      severity: json['severity'] as String? ?? 'error',
      messages: (json['messages'] as List<dynamic>? ?? const [])
          .map((message) => message.toString())
          .toList(),
    );
  }

  final String name;
  final String severity;
  final List<String> messages;
}

class DeploymentPage extends StatefulWidget {
  const DeploymentPage({
    required this.darkMode,
    required this.onToggleTheme,
    super.key,
  });

  final bool darkMode;
  final VoidCallback onToggleTheme;

  @override
  State<DeploymentPage> createState() => _DeploymentPageState();
}

class _DeploymentPageState extends State<DeploymentPage> {
  late final TextEditingController _projectRootController;
  final TextEditingController _pythonController = TextEditingController(
    text: 'python',
  );
  final TextEditingController _targetsController = TextEditingController();
  final TextEditingController _targetsFileController = TextEditingController();
  final TextEditingController _bgInfoFolderController = TextEditingController();
  final TextEditingController _workersController = TextEditingController(
    text: '10',
  );
  final ScrollController _outputScrollController = ScrollController();

  TargetSource _targetSource = TargetSource.file;
  bool _audioRecall = true;
  bool _displayRecall = true;
  bool _bgInfoInstall = false;
  bool _addDesktopShortcuts = true;
  bool _isRunning = false;
  bool _isMonitoring = false;
  bool _isUpdatingTargetsFile = false;
  bool _stopRequested = false;
  Process? _process;
  Process? _monitorProcess;
  StreamSubscription<String>? _stdoutSubscription;
  StreamSubscription<String>? _stderrSubscription;
  StreamSubscription<String>? _monitorStdoutSubscription;
  StreamSubscription<String>? _monitorStderrSubscription;
  String _output = '';
  String _status = 'Ready';
  String? _detailedLogPath;
  final Map<String, int> _detailedLogOffsets = {};
  List<String> _knownTargets = const [];
  Map<String, PcProgress> _pcProgress = const {};
  Set<String> _finishedTargets = const {};
  Map<String, List<ScriptDeploymentResult>> _scriptResultsByPc = const {};
  String _monitorStatus = 'Ready';
  List<String> _monitorTargets = const [];
  Set<String> _monitorCompletedTargets = const {};
  List<MonitoringPcResult> _monitorResults = const [];
  MonitoringFilter _monitorFilter = MonitoringFilter.all;

  @override
  void initState() {
    super.initState();
    _projectRootController = TextEditingController(text: _findProjectRoot());
    WidgetsBinding.instance.addPostFrameCallback((_) {
      _loadConfig(silent: true);
      _loadTargetsFile(silent: true);
    });
  }

  @override
  void dispose() {
    _stdoutSubscription?.cancel();
    _stderrSubscription?.cancel();
    _monitorStdoutSubscription?.cancel();
    _monitorStderrSubscription?.cancel();
    _projectRootController.dispose();
    _pythonController.dispose();
    _targetsController.dispose();
    _targetsFileController.dispose();
    _bgInfoFolderController.dispose();
    _workersController.dispose();
    _outputScrollController.dispose();
    super.dispose();
  }

  String _findProjectRoot() {
    final starts = <String>{
      Directory.current.absolute.path,
      File(Platform.resolvedExecutable).parent.absolute.path,
    };

    for (final start in starts) {
      var directory = Directory(start);
      for (var level = 0; level < 8; level++) {
        if (File(_join(directory.path, 'Deployment_Orchestrator.py'))
            .existsSync()) {
          return directory.path;
        }
        final parent = directory.parent;
        if (parent.path == directory.path) break;
        directory = parent;
      }
    }
    return Directory.current.absolute.path;
  }

  String _join(String parent, String child) =>
      '$parent${Platform.pathSeparator}$child';

  String _cleanExecutable(String value) {
    final trimmed = value.trim();
    if (trimmed.length >= 2 &&
        trimmed.startsWith('"') &&
        trimmed.endsWith('"')) {
      return trimmed.substring(1, trimmed.length - 1);
    }
    return trimmed;
  }

  bool? _readPythonBool(String source, String name) {
    final match = RegExp(
      '^\\s*${RegExp.escape(name)}\\s*=\\s*(True|False)',
      multiLine: true,
    ).firstMatch(source);
    return switch (match?.group(1)) {
      'True' => true,
      'False' => false,
      _ => null,
    };
  }

  String? _readPythonString(String source, String name) {
    final match = RegExp(
      '^\\s*${RegExp.escape(name)}\\s*=\\s*["\']([^"\']*)["\']',
      multiLine: true,
    ).firstMatch(source);
    return match?.group(1);
  }

  Future<void> _loadConfig({bool silent = false}) async {
    final root = _projectRootController.text.trim();
    final configFile = File(_join(root, 'config.py'));
    try {
      final source = await configFile.readAsString();
      if (!mounted) return;
      setState(() {
        final audioRecall =
            _readPythonBool(source, 'AUDIO_RECALL') ?? _audioRecall;
        final displayRecall =
            _readPythonBool(source, 'DISPLAY_RECALL') ?? _displayRecall;
        _audioRecall = audioRecall;
        _displayRecall = displayRecall;
        _bgInfoInstall =
            _readPythonBool(source, 'BGINFO_INSTALL') ?? _bgInfoInstall;
        final addDesktopShortcuts =
            _readPythonBool(source, 'ADD_DESKTOP_SHORTCUTS') ??
            _addDesktopShortcuts;
        _addDesktopShortcuts = audioRecall && displayRecall
            ? addDesktopShortcuts
            : false;
        _bgInfoFolderController.text =
            _readPythonString(source, 'BGINFO_FOLDER') ??
            _bgInfoFolderController.text;
      });
      if (!silent) _showMessage('Loaded settings from config.py.');
    } on FileSystemException {
      if (!silent) {
        _showMessage('Could not read ${configFile.path}.');
      }
    }
  }

  Future<void> _loadTargetsFile({bool silent = false}) async {
    final targetsFile = File(
      _join(_projectRootController.text.trim(), 'targets.txt'),
    );
    setState(() => _isUpdatingTargetsFile = true);
    try {
      final contents = await targetsFile.readAsString();
      if (!mounted) return;
      setState(() => _targetsFileController.text = contents);
      if (!silent) _showMessage('Loaded targets.txt.');
    } on FileSystemException {
      if (!silent) _showMessage('Could not read ${targetsFile.path}.');
    } finally {
      if (mounted) setState(() => _isUpdatingTargetsFile = false);
    }
  }

  Future<void> _saveTargetsFile() async {
    FocusManager.instance.primaryFocus?.unfocus();
    final targetsFile = File(
      _join(_projectRootController.text.trim(), 'targets.txt'),
    );
    setState(() => _isUpdatingTargetsFile = true);
    try {
      await targetsFile.writeAsString(_targetsFileController.text);
      if (!mounted) return;
      _showMessage('Saved targets.txt.');
    } on FileSystemException catch (error) {
      if (!mounted) return;
      _showMessage('Could not save targets.txt: ${error.message}');
    } finally {
      if (mounted) setState(() => _isUpdatingTargetsFile = false);
    }
  }

  List<String> _directTargets() {
    return _parseTargetText(_targetsController.text);
  }

  List<String> _parseTargetText(String contents) {
    return contents
        .split(RegExp(r'[\r\n,;]+'))
        .map((target) => target.trim())
        .where((target) => target.isNotEmpty && !target.startsWith('#'))
        .toSet()
        .toList();
  }

  Future<void> _loadDetailedLogFile({bool showErrors = true}) async {
    final path = _detailedLogPath;
    if (path == null || path.isEmpty) return;
    try {
      final contents = await File(path).readAsString();
      if (!mounted) return;
      final previousLength = _detailedLogOffsets[path] ?? 0;
      if (contents.length > previousLength) {
        final newContents = contents.substring(previousLength).trimRight();
        if (newContents.isNotEmpty) {
          final separator = previousLength == 0
              ? '\n\n===== Detailed log: $path =====\n'
              : '\n';
          setState(() {
            _output = _output.isEmpty
                ? '$separator$newContents'.trimLeft()
                : '$_output$separator$newContents';
          });
        }
      }
      _detailedLogOffsets[path] = contents.length;
    } on FileSystemException catch (error) {
      if (!mounted) return;
      if (showErrors) {
        _showMessage('Could not open detailed log: ${error.message}');
      }
    }
  }

  String _filteredOutputFor(String? selectedPc) {
    return filterDetailedLog(_output, selectedPc: selectedPc);
  }

  Future<void> _showDetailedLogModal({String? initialPc}) async {
    if (_detailedLogPath != null) {
      await _loadDetailedLogFile(showErrors: false);
    }
    if (!mounted) return;
    var selectedPc = initialPc;
    await showDialog<void>(
      context: context,
      builder: (dialogContext) => StatefulBuilder(
        builder: (dialogContext, setDialogState) {
          final filteredOutput = _filteredOutputFor(selectedPc);
          final selectedScriptResults = selectedPc == null
              ? const <ScriptDeploymentResult>[]
              : completeScriptStatusList(
                  _scriptResultsByPc[selectedPc] ??
                      const <ScriptDeploymentResult>[],
                );
          return AlertDialog(
            title: const Row(
              children: [
                Icon(Icons.article_outlined),
                SizedBox(width: 10),
                Text('Detailed log'),
              ],
            ),
            content: SizedBox(
              width: 940,
              height: (MediaQuery.sizeOf(dialogContext).height - 180).clamp(
                300,
                700,
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Autocomplete<String>(
                    initialValue: TextEditingValue(text: initialPc ?? ''),
                    displayStringForOption: (pc) => pc,
                    optionsBuilder: (value) =>
                        filterPcChoices(_knownTargets, value.text),
                    onSelected: (pc) {
                      setDialogState(() => selectedPc = pc);
                    },
                    fieldViewBuilder:
                        (context, controller, focusNode, onFieldSubmitted) {
                          return TextField(
                            key: const Key('detailedLogSearchField'),
                            controller: controller,
                            focusNode: focusNode,
                            autofocus: initialPc == null,
                            decoration: InputDecoration(
                              labelText: 'Filter by PC',
                              hintText: 'Start typing a PC name',
                              prefixIcon: const Icon(Icons.search),
                              suffixIcon: selectedPc == null
                                  ? null
                                  : IconButton(
                                      tooltip: 'Show all PCs',
                                      onPressed: () {
                                        controller.clear();
                                        setDialogState(() => selectedPc = null);
                                        focusNode.requestFocus();
                                      },
                                      icon: const Icon(Icons.clear),
                                    ),
                              border: const OutlineInputBorder(),
                              isDense: true,
                            ),
                            onChanged: (_) {
                              if (selectedPc != null) {
                                setDialogState(() => selectedPc = null);
                              }
                            },
                          );
                        },
                  ),
                  const SizedBox(height: 12),
                  Container(
                    key: const Key('scriptStatusPanel'),
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: Theme.of(dialogContext)
                          .colorScheme
                          .surfaceContainerHighest,
                      borderRadius: BorderRadius.circular(8),
                      border: Border.all(
                        color: Theme.of(dialogContext)
                            .colorScheme
                            .outlineVariant,
                      ),
                    ),
                    child: selectedPc == null
                        ? const Row(
                            children: [
                              Icon(Icons.info_outline, size: 20),
                              SizedBox(width: 8),
                              Text('Select a PC to view script statuses.'),
                            ],
                          )
                        : Column(
                            crossAxisAlignment: CrossAxisAlignment.stretch,
                            children: [
                              Text(
                                'Script statuses — $selectedPc',
                                style: Theme.of(dialogContext)
                                    .textTheme
                                    .labelLarge,
                              ),
                              const SizedBox(height: 8),
                              for (final script in selectedScriptResults)
                                Padding(
                                  padding: const EdgeInsets.only(top: 6),
                                  child: Row(
                                    children: [
                                      Icon(
                                        _severityIcon(script.severity),
                                        color: _severityColor(script.severity),
                                        size: 20,
                                      ),
                                      const SizedBox(width: 8),
                                      Expanded(
                                        child: Text(
                                          _scriptDisplayName(script.name),
                                          overflow: TextOverflow.ellipsis,
                                        ),
                                      ),
                                      const SizedBox(width: 12),
                                      Text(
                                        _resultLabel(script.severity),
                                        style: TextStyle(
                                          color: _severityColor(
                                            script.severity,
                                          ),
                                          fontWeight: FontWeight.w600,
                                        ),
                                      ),
                                    ],
                                  ),
                                ),
                            ],
                          ),
                  ),
                  const SizedBox(height: 12),
                  Expanded(
                    child: Container(
                      decoration: BoxDecoration(
                        color: const Color(0xff111827),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      padding: const EdgeInsets.all(12),
                      child: Scrollbar(
                        controller: _outputScrollController,
                        thumbVisibility: true,
                        child: SingleChildScrollView(
                          controller: _outputScrollController,
                          child: SelectableText.rich(
                            TextSpan(
                              style: const TextStyle(
                                color: Color(0xffd1d5db),
                                fontFamily: 'monospace',
                                fontSize: 13,
                              ),
                              children: buildSeveritySpans(
                                _output.isEmpty
                                    ? 'No detailed log output is available.'
                                    : filteredOutput.isEmpty
                                    ? 'No log entries for the selected PC.'
                                    : filteredOutput,
                              ),
                            ),
                            key: const Key('outputText'),
                          ),
                        ),
                      ),
                    ),
                  ),
                ],
              ),
            ),
            actions: [
              if (_detailedLogPath != null)
                IconButton(
                  tooltip: 'Reload detailed log',
                  onPressed: () async {
                    await _loadDetailedLogFile();
                    if (dialogContext.mounted) setDialogState(() {});
                  },
                  icon: const Icon(Icons.refresh),
                ),
              FilledButton(
                onPressed: () => Navigator.of(dialogContext).pop(),
                child: const Text('Close'),
              ),
            ],
          );
        },
      ),
    );
  }

  int _progressRank(PcProgress progress) {
    return switch (progress) {
      PcProgress.queued => 0,
      PcProgress.running => 1,
      PcProgress.complete => 2,
      PcProgress.offline => 3,
      PcProgress.warning => 4,
      PcProgress.error => 5,
      PcProgress.fatal => 6,
    };
  }

  void _updateProgressFromLog(String line) {
    final match = RegExp(
      r'\b(DEBUG|INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\s+([^:\s]+):\s*(.*)$',
      caseSensitive: false,
    ).firstMatch(line);
    if (match == null) return;
    final pc = match.group(2)!;
    if (!_pcProgress.containsKey(pc)) return;
    final severity = match.group(1)!.toUpperCase();
    final message = match.group(3)!;
    final next = message.startsWith('Ping test failed')
        ? PcProgress.offline
        : switch (severity) {
            'WARNING' || 'WARN' => PcProgress.warning,
            'ERROR' => PcProgress.error,
            'FATAL' || 'CRITICAL' => PcProgress.fatal,
            _ when message.startsWith('Queuing configuration check') =>
              PcProgress.queued,
            _ when message.startsWith('Deployment complete') =>
              PcProgress.complete,
            _ => PcProgress.running,
          };
    final current = _pcProgress[pc]!;
    final targetFinished = message.startsWith('Deployment complete');
    if (_progressRank(next) < _progressRank(current)) {
      if (targetFinished) {
        setState(() => _finishedTargets = {..._finishedTargets, pc});
      }
      return;
    }
    setState(() {
      _pcProgress = {..._pcProgress, pc: next};
      if (targetFinished) {
        _finishedTargets = {..._finishedTargets, pc};
      }
    });
  }

  Future<DeploymentReport?> _startDeployment() async {
    FocusManager.instance.primaryFocus?.unfocus();
    final root = _projectRootController.text.trim();
    final python = _cleanExecutable(_pythonController.text);
    final workers = int.tryParse(_workersController.text.trim());
    final orchestrator = File(_join(root, 'Deployment_Orchestrator.py'));
    final targetFile = File(_join(root, 'targets.txt'));
    final directTargets = _directTargets();

    if (!orchestrator.existsSync()) {
      _showMessage(
        'Deployment_Orchestrator.py was not found in the project root.',
      );
      return null;
    }
    if (python.isEmpty) {
      _showMessage('Enter a Python executable or command.');
      return null;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return null;
    }
    if (_bgInfoInstall && _bgInfoFolderController.text.trim().isEmpty) {
      _showMessage('Enter a BGInfo folder when BGInfo is enabled.');
      return null;
    }
    if (_targetSource == TargetSource.file && !targetFile.existsSync()) {
      _showMessage('targets.txt was not found in the project root.');
      return null;
    }
    if (_targetSource == TargetSource.direct && directTargets.isEmpty) {
      _showMessage('Enter at least one target PC.');
      return null;
    }

    final deploymentTargets = _targetSource == TargetSource.direct
        ? directTargets
        : _parseTargetText(targetFile.readAsStringSync());
    if (deploymentTargets.isEmpty) {
      _showMessage('No target PCs were provided.');
      return null;
    }
    final resultFile = _join(
      _join(root, 'logs'),
      'app-result-${DateTime.now().millisecondsSinceEpoch}.json',
    );

    final arguments = <String>[
      orchestrator.path,
      _audioRecall ? '--audio-recall' : '--no-audio-recall',
      _displayRecall ? '--display-recall' : '--no-display-recall',
      _bgInfoInstall ? '--bginfo-install' : '--no-bginfo-install',
      _addDesktopShortcuts
          ? '--add-desktop-shortcuts'
          : '--no-add-desktop-shortcuts',
      '--bginfo-folder',
      _bgInfoFolderController.text.trim(),
      '--max-workers',
      workers.toString(),
      '--result-file',
      resultFile,
    ];
    if (_targetSource == TargetSource.file) {
      arguments.addAll(['--targets-file', _join(root, 'targets.txt')]);
    } else {
      for (final target in directTargets) {
        arguments.addAll(['--target', target]);
      }
    }

    setState(() {
      _isRunning = true;
      _status = 'Starting deployment…';
      _detailedLogPath = null;
      _knownTargets = deploymentTargets;
      _pcProgress = {
        for (final target in deploymentTargets) target: PcProgress.queued,
      };
      _finishedTargets = const {};
      _stopRequested = false;
    });

    try {
      final process = await Process.start(
        python,
        arguments,
        workingDirectory: root,
        runInShell: false,
      );
      _process = process;
      if (!mounted) {
        process.kill();
        return null;
      }
      setState(() => _status = 'Deploying (PID ${process.pid})');

      _stdoutSubscription = process.stdout
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen((line) => _appendOutput(line));
      _stderrSubscription = process.stderr
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen((line) => _appendOutput('ERROR: $line'));

      final exitCode = await process.exitCode;
      await _stdoutSubscription?.cancel();
      await _stderrSubscription?.cancel();
      if (!mounted) return null;
      DeploymentReport? report;
      try {
        final reportJson = jsonDecode(await File(resultFile).readAsString());
        report = DeploymentReport.fromJson(reportJson as Map<String, dynamic>);
      } on FileSystemException catch (error) {
        _appendOutput(
          'ERROR: Could not read deployment summary: ${error.message}',
        );
      } on FormatException catch (error) {
        _appendOutput('ERROR: Invalid deployment summary: ${error.message}');
      }
      if (report != null) _appendReportScriptStatuses(report);
      final reportHasProblems =
          report?.pcs.any((result) => result.hasProblems) ?? exitCode != 0;
      setState(() {
        _process = null;
        _isRunning = false;
        _detailedLogPath = report?.logFile;
        if (report != null) {
          _knownTargets = report.pcs.map((result) => result.pc).toList();
          _pcProgress = {
            for (final result in report.pcs)
              result.pc: _progressForResult(result),
          };
          _finishedTargets = report.pcs.map((result) => result.pc).toSet();
        }
        _status = _stopRequested
            ? 'Deployment stopped'
            : exitCode != 0
            ? 'Deployment exited with code $exitCode'
            : reportHasProblems
            ? 'Deployment finished with issues'
            : 'Deployment finished successfully';
      });
      if (!_stopRequested) {
        report ??= DeploymentReport(
          logFile: '',
          pcs: deploymentTargets
              .map(
                (pc) => PcDeploymentResult(
                  pc: pc,
                  highestSeverity: 'fatal',
                  issues: const [
                    DeploymentIssue(
                      component: 'Orchestrator',
                      severity: 'fatal',
                      message:
                          'A structured deployment result was not available.',
                    ),
                  ],
                  scripts: const [],
                ),
              )
              .toList(),
        );
      }
      return report;
    } on ProcessException catch (error) {
      if (!mounted) return null;
      setState(() {
        _process = null;
        _isRunning = false;
        _status = 'Could not start deployment';
      });
      _appendOutput(error.message);
      return null;
    }
  }

  Future<DeploymentReport?> _startPostDeploymentRetry({
    required String target,
    required bool refreshDeploymentArea,
    ValueChanged<String>? onProgress,
  }) async {
    final root = _projectRootController.text.trim();
    final python = _cleanExecutable(_pythonController.text);
    final workers = int.tryParse(_workersController.text.trim());
    final orchestrator = File(_join(root, 'Deployment_Orchestrator.py'));
    if (!orchestrator.existsSync()) {
      _showMessage(
        'Deployment_Orchestrator.py was not found in the project root.',
      );
      return null;
    }
    if (python.isEmpty) {
      _showMessage('Enter a Python executable or command.');
      return null;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return null;
    }
    if (_bgInfoInstall && _bgInfoFolderController.text.trim().isEmpty) {
      _showMessage('Enter a BGInfo folder when BGInfo is enabled.');
      return null;
    }

    final resultFile = _join(
      _join(root, 'logs'),
      'app-retry-${DateTime.now().microsecondsSinceEpoch}.json',
    );
    final arguments = <String>[
      orchestrator.path,
      _audioRecall ? '--audio-recall' : '--no-audio-recall',
      _displayRecall ? '--display-recall' : '--no-display-recall',
      _bgInfoInstall ? '--bginfo-install' : '--no-bginfo-install',
      _addDesktopShortcuts
          ? '--add-desktop-shortcuts'
          : '--no-add-desktop-shortcuts',
      '--bginfo-folder',
      _bgInfoFolderController.text.trim(),
      '--max-workers',
      workers.toString(),
      '--result-file',
      resultFile,
      '--target',
      target,
    ];

    if (refreshDeploymentArea) {
      setState(() {
        _pcProgress = {..._pcProgress, target: PcProgress.queued};
        _finishedTargets = {..._finishedTargets.where((pc) => pc != target)};
      });
    }

    try {
      final process = await Process.start(
        python,
        arguments,
        workingDirectory: root,
        runInShell: false,
      );
      if (!mounted) {
        process.kill();
        return null;
      }
      final stdoutSubscription = process.stdout
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen(
            (line) => _appendOutput(
              line,
              onProgress: onProgress,
              updateProgress: refreshDeploymentArea,
            ),
          );
      final stderrSubscription = process.stderr
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen(
            (line) => _appendOutput(
              'ERROR: $line',
              onProgress: onProgress,
              updateProgress: refreshDeploymentArea,
            ),
          );
      await process.exitCode;
      await stdoutSubscription.cancel();
      await stderrSubscription.cancel();
      if (!mounted) return null;

      DeploymentReport? report;
      try {
        final reportJson = jsonDecode(await File(resultFile).readAsString());
        report = DeploymentReport.fromJson(reportJson as Map<String, dynamic>);
      } on FileSystemException catch (error) {
        _appendOutput(
          'ERROR: Could not read deployment summary: ${error.message}',
        );
      } on FormatException catch (error) {
        _appendOutput('ERROR: Invalid deployment summary: ${error.message}');
      }
      report ??= _failedRetryReport(target);
      _appendReportScriptStatuses(report);
      if (!mounted) return report;
      final targetResult = report.pcs.firstWhere(
        (result) => result.pc == target,
        orElse: () => _failedRetryReport(target).pcs.single,
      );
      setState(() {
        if (report!.logFile.isNotEmpty) _detailedLogPath = report.logFile;
        if (refreshDeploymentArea) {
          _pcProgress = {
            ..._pcProgress,
            target: _progressForResult(targetResult),
          };
          _finishedTargets = {..._finishedTargets, target};
        }
      });
      return report;
    } on ProcessException catch (error) {
      _appendOutput(error.message, onProgress: onProgress);
      if (refreshDeploymentArea && mounted) {
        final failedResult = _failedRetryReport(target).pcs.single;
        setState(() {
          _pcProgress = {
            ..._pcProgress,
            target: _progressForResult(failedResult),
          };
          _finishedTargets = {..._finishedTargets, target};
        });
      }
      return null;
    }
  }

  Future<void> _startMonitoring() async {
    FocusManager.instance.primaryFocus?.unfocus();
    final root = _projectRootController.text.trim();
    final python = _cleanExecutable(_pythonController.text);
    final workers = int.tryParse(_workersController.text.trim());
    final orchestrator = File(_join(root, 'Monitoring_Orchestrator.py'));
    final targetFile = File(_join(root, 'targets.txt'));
    final directTargets = _directTargets();

    if (!orchestrator.existsSync()) {
      _showMessage(
        'Monitoring_Orchestrator.py was not found in the project root.',
      );
      return;
    }
    if (python.isEmpty) {
      _showMessage('Enter a Python executable or command.');
      return;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return;
    }
    if (_targetSource == TargetSource.file && !targetFile.existsSync()) {
      _showMessage('targets.txt was not found in the project root.');
      return;
    }
    if (_targetSource == TargetSource.direct && directTargets.isEmpty) {
      _showMessage('Enter at least one target PC on the Deployment tab.');
      return;
    }

    final targets = _targetSource == TargetSource.direct
        ? directTargets
        : _parseTargetText(targetFile.readAsStringSync());
    if (targets.isEmpty) {
      _showMessage('No target PCs were provided.');
      return;
    }
    final resultFile = _join(
      _join(root, 'logs'),
      'monitor-result-${DateTime.now().millisecondsSinceEpoch}.json',
    );
    final arguments = <String>[
      orchestrator.path,
      '--max-workers',
      workers.toString(),
      '--result-file',
      resultFile,
    ];
    if (_targetSource == TargetSource.file) {
      arguments.addAll(['--targets-file', targetFile.path]);
    } else {
      for (final target in targets) {
        arguments.addAll(['--target', target]);
      }
    }

    setState(() {
      _isMonitoring = true;
      _monitorStatus = 'Monitoring ${targets.length} target(s)…';
      _monitorTargets = targets;
      _monitorCompletedTargets = const {};
      _monitorResults = const [];
      _monitorFilter = MonitoringFilter.all;
    });

    try {
      final process = await Process.start(
        python,
        arguments,
        workingDirectory: root,
        runInShell: false,
      );
      _monitorProcess = process;
      if (!mounted) {
        process.kill();
        return;
      }
      _monitorStdoutSubscription = process.stdout
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen((line) {
            if (!mounted) return;
            const marker = 'MONITOR_PROGRESS ';
            if (!line.startsWith(marker)) return;
            final pc = line.substring(marker.length).trim();
            if (pc.isEmpty) return;
            setState(() {
              _monitorCompletedTargets = {..._monitorCompletedTargets, pc};
            });
          });
      _monitorStderrSubscription = process.stderr
          .transform(utf8.decoder)
          .transform(const LineSplitter())
          .listen((_) {});
      final exitCode = await process.exitCode;
      await _monitorStdoutSubscription?.cancel();
      await _monitorStderrSubscription?.cancel();
      if (!mounted) return;

      MonitoringReport? report;
      try {
        final reportJson = jsonDecode(await File(resultFile).readAsString());
        report = MonitoringReport.fromJson(reportJson as Map<String, dynamic>);
      } on FileSystemException catch (error) {
        _showMessage('Could not read monitoring results: ${error.message}');
      } on FormatException catch (error) {
        _showMessage('Invalid monitoring results: ${error.message}');
      }
      setState(() {
        _monitorProcess = null;
        _isMonitoring = false;
        _monitorResults = report?.pcs ?? const [];
        if (report != null) {
          _monitorCompletedTargets = _monitorTargets.toSet();
        }
        _monitorStatus = exitCode == 0 && report != null
            ? 'Monitoring complete'
            : 'Monitoring exited with code $exitCode';
      });
    } on ProcessException catch (error) {
      if (!mounted) return;
      setState(() {
        _monitorProcess = null;
        _isMonitoring = false;
        _monitorStatus = 'Could not start monitoring';
      });
      _showMessage(error.message);
    }
  }

  void _stopMonitoring() {
    if (_monitorProcess?.kill() ?? false) {
      setState(() => _monitorStatus = 'Stopping monitoring…');
    } else {
      _showMessage('The monitoring process could not be stopped.');
    }
  }

  DeploymentReport _failedRetryReport(String target) {
    return DeploymentReport(
      logFile: '',
      pcs: [
        PcDeploymentResult(
          pc: target,
          highestSeverity: 'fatal',
          issues: const [
            DeploymentIssue(
              component: 'Orchestrator',
              severity: 'fatal',
              message: 'A structured deployment result was not available.',
            ),
          ],
          scripts: const [],
        ),
      ],
    );
  }

  void _appendOutput(
    String line, {
    ValueChanged<String>? onProgress,
    bool updateProgress = true,
  }) {
    if (!mounted) return;
    setState(() {
      _output = _output.isEmpty ? line : '$_output\n$line';
    });
    if (updateProgress) _updateProgressFromLog(line);
    onProgress?.call(line);
    WidgetsBinding.instance.addPostFrameCallback((_) {
      if (_outputScrollController.hasClients) {
        _outputScrollController.jumpTo(
          _outputScrollController.position.maxScrollExtent,
        );
      }
    });
  }

  PcProgress _progressForResult(PcDeploymentResult result) {
    if (result.issues.any(
      (issue) =>
          issue.component == 'Connectivity' &&
          issue.message.startsWith('Ping test failed'),
    )) {
      return PcProgress.offline;
    }
    return switch (result.highestSeverity.toLowerCase()) {
      'warning' => PcProgress.warning,
      'error' => PcProgress.error,
      'fatal' => PcProgress.fatal,
      _ => PcProgress.complete,
    };
  }

  void _appendReportScriptStatuses(DeploymentReport report) {
    setState(() {
      _scriptResultsByPc = {
        ..._scriptResultsByPc,
        for (final pcResult in report.pcs) pcResult.pc: pcResult.scripts,
      };
    });
    for (final pcResult in report.pcs) {
      for (final script in pcResult.scripts) {
        final scriptName = script.name.split(RegExp(r'[\\/]')).last;
        final severity = script.severity.toUpperCase();
        final detail = script.messages.isEmpty
            ? ''
            : ' — ${script.messages.join(' | ')}';
        _appendOutput(
          '$severity ${pcResult.pc}: Script status: $scriptName ($severity)$detail',
        );
      }
    }
  }

  void _stopDeployment() {
    final stopped = _process?.kill() ?? false;
    if (stopped) {
      setState(() {
        _stopRequested = true;
        _status = 'Stopping deployment…';
      });
    } else {
      _showMessage('The deployment process could not be stopped.');
    }
  }

  Color _severityColor(String severity) {
    return switch (severity.toLowerCase()) {
      'info' => const Color(0xff22c55e),
      'warning' => const Color(0xffff9800),
      'error' || 'fatal' => const Color(0xffef4444),
      'not_deployed' => const Color(0xff94a3b8),
      _ => Theme.of(context).colorScheme.onSurfaceVariant,
    };
  }

  IconData _severityIcon(String severity) {
    return switch (severity.toLowerCase()) {
      'info' => Icons.check_circle,
      'warning' => Icons.warning_amber_rounded,
      'error' || 'fatal' => Icons.error,
      'not_deployed' => Icons.block,
      _ => Icons.help,
    };
  }

  String _scriptDisplayName(String path) {
    return scriptFileName(path);
  }

  String _resultLabel(String severity) {
    return switch (severity.toLowerCase()) {
      'info' => 'Complete',
      'warning' => 'Warning',
      'error' => 'Error',
      'fatal' => 'Fatal',
      'not_deployed' => 'Not deployed',
      _ => severity,
    };
  }

  Future<void> _showDeploymentSummary(DeploymentReport report) async {
    if (!mounted) return;
    final results = [...report.pcs];
    var detailedLogPath = report.logFile;
    final retryingPcs = <String>{};
    final retryProgressByPc = <String, String>{};
    await showDialog<void>(
      context: context,
      barrierDismissible: false,
      builder: (dialogContext) {
        return StatefulBuilder(
          builder: (dialogContext, setDialogState) {
            final availableHeight =
                MediaQuery.sizeOf(dialogContext).height - 180;
            Future<void> retry(PcDeploymentResult result) async {
              if (retryingPcs.contains(result.pc)) return;
              setDialogState(() {
                retryingPcs.add(result.pc);
                retryProgressByPc[result.pc] = 'Starting retry…';
              });
              final retryReport = await _startPostDeploymentRetry(
                target: result.pc,
                refreshDeploymentArea: false,
                onProgress: (line) {
                  if (!dialogContext.mounted) return;
                  setDialogState(() => retryProgressByPc[result.pc] = line);
                },
              );
              if (!dialogContext.mounted) return;
              PcDeploymentResult? replacement;
              if (retryReport != null) {
                for (final item in retryReport.pcs) {
                  if (item.pc == result.pc) {
                    replacement = item;
                    break;
                  }
                }
              }
              setDialogState(() {
                if (retryReport != null && retryReport.logFile.isNotEmpty) {
                  detailedLogPath = retryReport.logFile;
                }
                results[results.indexWhere((item) => item.pc == result.pc)] =
                    replacement ??
                    PcDeploymentResult(
                      pc: result.pc,
                      highestSeverity: 'fatal',
                      issues: const [
                        DeploymentIssue(
                          component: 'Orchestrator',
                          severity: 'fatal',
                          message: 'The retry did not return a result.',
                        ),
                      ],
                      scripts: const [],
                    );
                retryingPcs.remove(result.pc);
                retryProgressByPc.remove(result.pc);
              });
            }

            return AlertDialog(
              title: const Row(
                children: [
                  Icon(Icons.fact_check_outlined),
                  SizedBox(width: 10),
                  Text('Deployment results'),
                ],
              ),
              content: SizedBox(
                width: 720,
                height: availableHeight.clamp(260, 650),
                child: results.isEmpty
                    ? const Center(child: Text('No PC results were returned.'))
                    : ListView.separated(
                        itemCount: results.length,
                        separatorBuilder: (_, _) => const SizedBox(height: 10),
                        itemBuilder: (_, index) => _buildPcResultCard(
                          results[index],
                          retrying: retryingPcs.contains(results[index].pc),
                          retryEnabled: !retryingPcs.contains(
                            results[index].pc,
                          ),
                          retryProgress:
                              retryProgressByPc[results[index].pc] ?? '',
                          onRetry: () => retry(results[index]),
                        ),
                      ),
              ),
              actions: [
                if (detailedLogPath.isNotEmpty)
                  TextButton.icon(
                    key: const Key('summaryDetailedLogButton'),
                    onPressed: () {
                      _detailedLogPath = detailedLogPath;
                      _showDetailedLogModal();
                    },
                    icon: const Icon(Icons.article_outlined),
                    label: const Text('View detailed log'),
                  ),
                FilledButton(
                  onPressed: () => Navigator.of(dialogContext).pop(),
                  child: const Text('Close'),
                ),
              ],
            );
          },
        );
      },
    );
  }

  Widget _buildPcResultCard(
    PcDeploymentResult result, {
    required bool retrying,
    required bool retryEnabled,
    required String retryProgress,
    required VoidCallback onRetry,
  }) {
    if (!result.hasProblems) {
      return Card(
        color: const Color(0xff22c55e).withValues(alpha: 0.12),
        child: ListTile(
          leading: const Icon(Icons.check_circle, color: Color(0xff22c55e)),
          title: Text(result.pc),
          subtitle: const Text('Complete'),
        ),
      );
    }

    final severityColor = _severityColor(result.highestSeverity);
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              children: [
                Icon(
                  _severityIcon(result.highestSeverity),
                  color: severityColor,
                ),
                const SizedBox(width: 10),
                Expanded(
                  child: Text(
                    result.pc,
                    style: Theme.of(context).textTheme.titleMedium,
                  ),
                ),
                Text(
                  _resultLabel(result.highestSeverity),
                  style: TextStyle(
                    color: severityColor,
                    fontWeight: result.highestSeverity == 'fatal'
                        ? FontWeight.bold
                        : FontWeight.w600,
                  ),
                ),
              ],
            ),
            if (result.issues.isNotEmpty) ...[
              const SizedBox(height: 12),
              Text('Issues', style: Theme.of(context).textTheme.labelLarge),
              for (final issue in result.issues)
                Padding(
                  padding: const EdgeInsets.only(top: 6),
                  child: Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Icon(
                        _severityIcon(issue.severity),
                        size: 18,
                        color: _severityColor(issue.severity),
                      ),
                      const SizedBox(width: 8),
                      Expanded(
                        child: Text(
                          '${issue.component}: ${issue.message}',
                          style: TextStyle(
                            fontWeight: issue.severity == 'fatal'
                                ? FontWeight.bold
                                : null,
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
            ],
            const SizedBox(height: 12),
            Text('Scripts run', style: Theme.of(context).textTheme.labelLarge),
            if (result.scripts.isEmpty)
              const Padding(
                padding: EdgeInsets.only(top: 6),
                child: Text('No deployment scripts ran.'),
              )
            else
              for (final script in result.scripts)
                ListTile(
                  dense: true,
                  contentPadding: EdgeInsets.zero,
                  leading: Icon(
                    _severityIcon(script.severity),
                    color: _severityColor(script.severity),
                  ),
                  title: Text(_scriptDisplayName(script.name)),
                  subtitle: script.messages.isEmpty
                      ? null
                      : Text(script.messages.join('\n')),
                  trailing: Text(
                    _resultLabel(script.severity),
                    style: TextStyle(
                      color: _severityColor(script.severity),
                      fontWeight: script.severity == 'fatal'
                          ? FontWeight.bold
                          : FontWeight.w600,
                    ),
                  ),
                ),
            if (retrying) ...[
              const SizedBox(height: 8),
              const LinearProgressIndicator(),
              const SizedBox(height: 8),
              Text(
                retryProgress,
                key: Key('retryProgress-${result.pc}'),
                maxLines: 2,
                overflow: TextOverflow.ellipsis,
              ),
            ],
            Align(
              alignment: Alignment.centerRight,
              child: OutlinedButton.icon(
                key: Key('retryButton-${result.pc}'),
                onPressed: retryEnabled ? onRetry : null,
                icon: const Icon(Icons.refresh),
                label: Text(
                  retrying ? 'Retrying ${result.pc}' : 'Retry ${result.pc}',
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  void _showMessage(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(context)
      ..hideCurrentSnackBar()
      ..showSnackBar(SnackBar(content: Text(message)));
  }

  @override
  Widget build(BuildContext context) {
    return DefaultTabController(
      length: 2,
      child: Scaffold(
        appBar: AppBar(
          title: const Text('Deployment Orchestrator'),
          backgroundColor: Theme.of(context)
              .colorScheme
              .surfaceContainerHighest,
          actions: [
            TextButton.icon(
              key: const Key('settingsButton'),
              onPressed: _showSettingsDialog,
              icon: const Icon(Icons.settings_outlined),
              label: const Text('Settings'),
            ),
            IconButton(
              key: const Key('themeToggleButton'),
              onPressed: widget.onToggleTheme,
              tooltip: widget.darkMode
                  ? 'Switch to light mode'
                  : 'Switch to dark mode',
              icon: Icon(widget.darkMode ? Icons.light_mode : Icons.dark_mode),
            ),
            const SizedBox(width: 8),
          ],
          bottom: const TabBar(
            tabs: [
              Tab(icon: Icon(Icons.rocket_launch_outlined), text: 'Deployment'),
              Tab(icon: Icon(Icons.monitor_heart_outlined), text: 'Monitoring'),
            ],
          ),
        ),
        body: SafeArea(
          child: TabBarView(
            children: [_buildDeploymentView(), _buildMonitoringView()],
          ),
        ),
      ),
    );
  }

  Widget _buildDeploymentView() {
    return SingleChildScrollView(
      padding: const EdgeInsets.all(20),
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 1000),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              LayoutBuilder(
                builder: (context, constraints) {
                  final narrow = constraints.maxWidth < 760;
                  final targetCard = _buildTargetsCard();
                  final featureCard = _buildFeaturesCard();
                  if (narrow) {
                    return Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        targetCard,
                        const SizedBox(height: 16),
                        featureCard,
                      ],
                    );
                  }
                  return Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Expanded(child: targetCard),
                      const SizedBox(width: 16),
                      Expanded(child: featureCard),
                    ],
                  );
                },
              ),
              const SizedBox(height: 16),
              _buildRunCard(),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildMonitoringView() {
    final targetDescription = _targetSource == TargetSource.file
        ? 'Using targets.txt'
        : 'Using ${_directTargets().length} directly entered target(s)';
    final filteredResults = _monitorResults
        .where((result) => _matchesMonitoringFilter(result, _monitorFilter))
        .toList();
    final monitoringProgress = _monitorTargets.isEmpty
        ? 0.0
        : _monitorCompletedTargets.length / _monitorTargets.length;
    return SingleChildScrollView(
      padding: const EdgeInsets.all(20),
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 1000),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Card(
                child: Padding(
                  padding: const EdgeInsets.all(16),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      Row(
                        children: [
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  'CTS deployment monitoring',
                                  style: Theme.of(context).textTheme.titleLarge,
                                ),
                                const SizedBox(height: 4),
                                Text('$targetDescription · $_monitorStatus'),
                              ],
                            ),
                          ),
                          if (_isMonitoring)
                            OutlinedButton.icon(
                              key: const Key('stopMonitoringButton'),
                              onPressed: _stopMonitoring,
                              icon: const Icon(Icons.stop),
                              label: const Text('Stop'),
                            )
                          else
                            FilledButton.icon(
                              key: const Key('startMonitoringButton'),
                              onPressed: _isRunning ? null : _startMonitoring,
                              icon: const Icon(Icons.refresh),
                              label: const Text('Run monitoring'),
                            ),
                        ],
                      ),
                      const SizedBox(height: 8),
                      const Text(
                        'Target selection is shared with the Deployment tab. Monitoring runs read-only PowerShell checks on each PC.',
                      ),
                      if (_isMonitoring) ...[
                        const SizedBox(height: 14),
                        LinearProgressIndicator(value: monitoringProgress),
                        const SizedBox(height: 8),
                        Text(
                          '${_monitorCompletedTargets.length} of ${_monitorTargets.length} targets inspected',
                          key: const Key('monitoringProgressText'),
                        ),
                      ],
                    ],
                  ),
                ),
              ),
              if (_monitorResults.isNotEmpty) ...[
                const SizedBox(height: 16),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(16),
                    child: Row(
                      crossAxisAlignment: CrossAxisAlignment.end,
                      children: [
                        Expanded(
                          child: DropdownButtonFormField<MonitoringFilter>(
                            key: const Key('monitoringReportFilter'),
                            initialValue: _monitorFilter,
                            decoration: const InputDecoration(
                              labelText: 'Monitoring report',
                              border: OutlineInputBorder(),
                              isDense: true,
                            ),
                            items: [
                              for (final filter in MonitoringFilter.values)
                                DropdownMenuItem(
                                  value: filter,
                                  child: Text(_monitoringFilterLabel(filter)),
                                ),
                            ],
                            onChanged: (filter) {
                              if (filter != null) {
                                setState(() => _monitorFilter = filter);
                              }
                            },
                          ),
                        ),
                        const SizedBox(width: 12),
                        OutlinedButton.icon(
                          key: const Key('copyMonitoringReportButton'),
                          onPressed: filteredResults.isEmpty
                              ? null
                              : () => _copyMonitoringReport(filteredResults),
                          icon: const Icon(Icons.copy),
                          label: const Text('Copy PC list'),
                        ),
                      ],
                    ),
                  ),
                ),
                const SizedBox(height: 10),
                Text(
                  '${_monitoringFilterLabel(_monitorFilter)} · ${filteredResults.length} PC(s)',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 10),
                if (filteredResults.isEmpty)
                  const Card(
                    child: Padding(
                      padding: EdgeInsets.all(20),
                      child: Center(
                        child: Text('No PCs match this monitoring report.'),
                      ),
                    ),
                  ),
                LayoutBuilder(
                  builder: (context, constraints) {
                    const spacing = 10.0;
                    final itemWidth =
                        (constraints.maxWidth - (spacing * 3)) / 4;
                    return Wrap(
                      spacing: spacing,
                      runSpacing: spacing,
                      children: [
                        for (final result in filteredResults)
                          SizedBox(
                            width: itemWidth,
                            child: _buildMonitoringResultTile(result),
                          ),
                      ],
                    );
                  },
                ),
              ] else if (!_isMonitoring) ...[
                const SizedBox(height: 24),
                const Center(
                  child: Text('Run monitoring to inspect the selected PCs.'),
                ),
              ],
            ],
          ),
        ),
      ),
    );
  }

  String _monitoringFilterLabel(MonitoringFilter filter) {
    return switch (filter) {
      MonitoringFilter.all => 'All PCs',
      MonitoringFilter.failedChecks => 'All PCs With Failed Checks',
      MonitoringFilter.offline => 'All Offline PCs',
      MonitoringFilter.winRmUnavailable => 'All PCs With WinRM Unavailable',
      MonitoringFilter.missingDeployment => 'All PCs Missing CTS Deployment',
      MonitoringFilter.missingAudioConfiguration =>
        'All PCs Missing Audio Configuration',
      MonitoringFilter.missingDisplayConfiguration =>
        'All PCs Missing Display Configuration',
      MonitoringFilter.missingLogoutShortcut =>
        'All PCs Missing Log Out Shortcut',
      MonitoringFilter.missingRebootShortcut =>
        'All PCs Missing Reboot Shortcut',
      MonitoringFilter.missingBgInfo => 'All PCs Missing BGInfo Deployment',
      MonitoringFilter.missingAudioDeviceCmdlets =>
        'All PCs Missing AudioDeviceCmdlets',
      MonitoringFilter.missingDisplayConfig => 'All PCs Missing DisplayConfig',
    };
  }

  bool _matchesMonitoringFilter(
    MonitoringPcResult result,
    MonitoringFilter filter,
  ) {
    final intent = result.deploymentIntent;
    final deploymentStatus = _monitoringComponentStatus(
      result,
      present: result.ctsDeployed,
      intended: true,
    );
    final audioStatus = _monitoringComponentStatus(
      result,
      present: result.audioConfigured,
      intended: intent?.audioRecall,
    );
    final displayStatus = _monitoringComponentStatus(
      result,
      present: result.displayConfigured,
      intended: intent?.displayRecall,
    );
    final logoutStatus = _monitoringComponentStatus(
      result,
      present: result.logoutShortcut,
      intended: intent?.desktopShortcuts,
    );
    final rebootStatus = _monitoringComponentStatus(
      result,
      present: result.rebootShortcut,
      intended: intent?.desktopShortcuts,
    );
    final bgInfoStatus = _monitoringComponentStatus(
      result,
      present: result.bgInfoDeployed,
      intended: intent?.bgInfoInstall,
    );
    final audioModuleStatus = _monitoringComponentStatus(
      result,
      present: result.audioDeviceCmdletsVersions.isNotEmpty,
      intended: intent?.audioRecall,
    );
    final displayModuleStatus = _monitoringComponentStatus(
      result,
      present: result.displayConfigVersions.isNotEmpty,
      intended: intent?.displayRecall,
    );
    final componentStatuses = [
      deploymentStatus,
      audioStatus,
      displayStatus,
      logoutStatus,
      rebootStatus,
      bgInfoStatus,
      audioModuleStatus,
      displayModuleStatus,
    ];
    final hasFailedCheck =
        !result.online ||
        !result.winRm ||
        componentStatuses.contains(MonitoringComponentStatus.missing);
    return switch (filter) {
      MonitoringFilter.all => true,
      MonitoringFilter.failedChecks => hasFailedCheck,
      MonitoringFilter.offline => !result.online,
      MonitoringFilter.winRmUnavailable => result.online && !result.winRm,
      MonitoringFilter.missingDeployment =>
        deploymentStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingAudioConfiguration =>
        audioStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingDisplayConfiguration =>
        displayStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingLogoutShortcut =>
        logoutStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingRebootShortcut =>
        rebootStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingBgInfo =>
        bgInfoStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingAudioDeviceCmdlets =>
        audioModuleStatus == MonitoringComponentStatus.missing,
      MonitoringFilter.missingDisplayConfig =>
        displayModuleStatus == MonitoringComponentStatus.missing,
    };
  }

  MonitoringComponentStatus _monitoringComponentStatus(
    MonitoringPcResult result, {
    required bool present,
    bool? intended,
  }) {
    if (!result.online || !result.winRm) {
      return MonitoringComponentStatus.unknown;
    }
    if (present) return MonitoringComponentStatus.present;
    if (intended == false) return MonitoringComponentStatus.notDeployed;
    return MonitoringComponentStatus.missing;
  }

  Future<void> _copyMonitoringReport(List<MonitoringPcResult> results) async {
    await Clipboard.setData(
      ClipboardData(text: results.map((result) => result.pc).join('\r\n')),
    );
    if (!mounted) return;
    _showMessage(
      'Copied ${results.length} PC(s) from ${_monitoringFilterLabel(_monitorFilter)}.',
    );
  }

  Widget _buildMonitoringResultTile(MonitoringPcResult result) {
    final Color overallColor;
    final IconData overallIcon;
    final String overallLabel;
    if (!result.online) {
      overallColor = const Color(0xff3b82f6);
      overallIcon = Icons.cloud_off;
      overallLabel = 'Offline';
    } else if (!result.winRm) {
      overallColor = const Color(0xffef4444);
      overallIcon = Icons.error;
      overallLabel = 'WinRM unavailable';
    } else if (!result.ctsDeployed) {
      overallColor = const Color(0xffff9800);
      overallIcon = Icons.warning_amber_rounded;
      overallLabel = 'Missing deployment';
    } else {
      final fullyConfigured = !_matchesMonitoringFilter(
        result,
        MonitoringFilter.failedChecks,
      );
      overallColor = fullyConfigured
          ? const Color(0xff22c55e)
          : const Color(0xffff9800);
      overallIcon = fullyConfigured
          ? Icons.check_circle
          : Icons.warning_amber_rounded;
      overallLabel = fullyConfigured ? 'Healthy' : 'Attention needed';
    }

    return Card(
      key: Key('monitorResult-${result.pc}'),
      clipBehavior: Clip.antiAlias,
      child: InkWell(
        onTap: () => _showMonitoringDetails(result),
        child: SizedBox(
          height: 78,
          child: Padding(
            padding: const EdgeInsets.symmetric(horizontal: 10),
            child: Row(
              children: [
                Icon(overallIcon, color: overallColor, size: 22),
                const SizedBox(width: 8),
                Expanded(
                  child: Tooltip(
                    message: result.pc,
                    child: Column(
                      mainAxisAlignment: MainAxisAlignment.center,
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          truncateComputerName(result.pc),
                          maxLines: 1,
                          overflow: TextOverflow.clip,
                          style: const TextStyle(fontWeight: FontWeight.w600),
                        ),
                        Text(
                          overallLabel,
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                          style: TextStyle(color: overallColor, fontSize: 12),
                        ),
                      ],
                    ),
                  ),
                ),
                const Icon(Icons.open_in_new, size: 16),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Future<void> _showMonitoringDetails(MonitoringPcResult result) async {
    final intent = result.deploymentIntent;
    final ctsStatus = _monitoringComponentStatus(
      result,
      present: result.ctsDeployed,
      intended: true,
    );
    final audioStatus = _monitoringComponentStatus(
      result,
      present: result.audioConfigured,
      intended: intent?.audioRecall,
    );
    final displayStatus = _monitoringComponentStatus(
      result,
      present: result.displayConfigured,
      intended: intent?.displayRecall,
    );
    final shortcutsIntended = intent?.desktopShortcuts;
    final logoutStatus = _monitoringComponentStatus(
      result,
      present: result.logoutShortcut,
      intended: shortcutsIntended,
    );
    final rebootStatus = _monitoringComponentStatus(
      result,
      present: result.rebootShortcut,
      intended: shortcutsIntended,
    );
    final bgInfoStatus = _monitoringComponentStatus(
      result,
      present: result.bgInfoDeployed,
      intended: intent?.bgInfoInstall,
    );
    final audioModuleStatus = _monitoringComponentStatus(
      result,
      present: result.audioDeviceCmdletsVersions.isNotEmpty,
      intended: intent?.audioRecall,
    );
    final displayModuleStatus = _monitoringComponentStatus(
      result,
      present: result.displayConfigVersions.isNotEmpty,
      intended: intent?.displayRecall,
    );
    final bgInfoDetail = bgInfoStatus == MonitoringComponentStatus.unknown
        ? 'Live details unavailable.'
        : bgInfoStatus == MonitoringComponentStatus.notDeployed
        ? 'Not requested by the latest deployment.'
        : 'EXE ${_yesNo(result.bgInfoExecutable)} · profile ${_yesNo(result.bgInfoProfile)} · background ${_yesNo(result.bgInfoBackground)} · startup ${_yesNo(result.bgInfoStartup)}${result.bgInfoStartupMethod.isEmpty ? '' : ' (${result.bgInfoStartupMethod})'}';

    await showDialog<void>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: Row(
          children: [
            const Icon(Icons.monitor_heart_outlined),
            const SizedBox(width: 10),
            Expanded(child: Text(result.pc)),
          ],
        ),
        content: SizedBox(
          width: 650,
          child: SingleChildScrollView(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                if (intent != null && intent.recordedAt.isNotEmpty) ...[
                  Text(
                    'Latest recorded deployment: ${intent.recordedAt}',
                    style: Theme.of(dialogContext).textTheme.bodySmall,
                  ),
                  const SizedBox(height: 12),
                ],
                Wrap(
                  spacing: 12,
                  runSpacing: 10,
                  children: [
                    _buildMonitorCheck('CTS deployment', ctsStatus),
                    _buildMonitorCheck('Audio configuration', audioStatus),
                    _buildMonitorCheck('Display configuration', displayStatus),
                    _buildMonitorCheck('Log Out.lnk', logoutStatus),
                    _buildMonitorCheck('Reboot.lnk', rebootStatus),
                    _buildMonitorCheck(
                      'BGInfo deployment',
                      bgInfoStatus,
                      detail: bgInfoDetail,
                    ),
                    _buildMonitorVersions(
                      'AudioDeviceCmdlets',
                      audioModuleStatus,
                      result.audioDeviceCmdletsVersions,
                    ),
                    _buildMonitorVersions(
                      'DisplayConfig',
                      displayModuleStatus,
                      result.displayConfigVersions,
                    ),
                  ],
                ),
              ],
            ),
          ),
        ),
        actions: [
          FilledButton(
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('Close'),
          ),
        ],
      ),
    );
  }

  String _yesNo(bool value) => value ? 'yes' : 'no';

  Widget _buildMonitorCheck(
    String label,
    MonitoringComponentStatus status, {
    String? detail,
  }) {
    final color = _monitoringStatusColor(status);
    return Container(
      width: 285,
      padding: const EdgeInsets.all(10),
      decoration: BoxDecoration(
        border: Border.all(color: color.withValues(alpha: 0.55)),
        borderRadius: BorderRadius.circular(8),
      ),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Icon(_monitoringStatusIcon(status), color: color),
          const SizedBox(width: 8),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  label,
                  style: const TextStyle(fontWeight: FontWeight.w600),
                ),
                Text(_monitoringStatusLabel(status)),
                if (detail != null)
                  Text(detail, style: Theme.of(context).textTheme.bodySmall),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildMonitorVersions(
    String module,
    MonitoringComponentStatus status,
    List<String> versions,
  ) {
    final color = _monitoringStatusColor(status);
    return Container(
      width: 285,
      padding: const EdgeInsets.all(10),
      decoration: BoxDecoration(
        border: Border.all(color: color.withValues(alpha: 0.55)),
        borderRadius: BorderRadius.circular(8),
      ),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Icon(_monitoringStatusIcon(status), color: color),
          const SizedBox(width: 8),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  module,
                  style: const TextStyle(fontWeight: FontWeight.w600),
                ),
                Text(
                  status == MonitoringComponentStatus.present
                      ? versions.join(', ')
                      : _monitoringStatusLabel(status),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Color _monitoringStatusColor(MonitoringComponentStatus status) {
    return switch (status) {
      MonitoringComponentStatus.present => const Color(0xff22c55e),
      MonitoringComponentStatus.missing => const Color(0xffef4444),
      MonitoringComponentStatus.notDeployed => const Color(0xff94a3b8),
      MonitoringComponentStatus.unknown => const Color(0xff3b82f6),
    };
  }

  IconData _monitoringStatusIcon(MonitoringComponentStatus status) {
    return switch (status) {
      MonitoringComponentStatus.present => Icons.check_circle,
      MonitoringComponentStatus.missing => Icons.cancel,
      MonitoringComponentStatus.notDeployed => Icons.block,
      MonitoringComponentStatus.unknown => Icons.help_outline,
    };
  }

  String _monitoringStatusLabel(MonitoringComponentStatus status) {
    return switch (status) {
      MonitoringComponentStatus.present => 'Present',
      MonitoringComponentStatus.missing => 'Missing',
      MonitoringComponentStatus.notDeployed => 'Not Deployed',
      MonitoringComponentStatus.unknown => 'Unknown',
    };
  }

  Future<void> _showSettingsDialog() async {
    await showDialog<void>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.settings_outlined),
            SizedBox(width: 10),
            Text('Settings'),
          ],
        ),
        content: SizedBox(
          width: 650,
          child: SingleChildScrollView(child: _buildRuntimeCard()),
        ),
        actions: [
          FilledButton(
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('Done'),
          ),
        ],
      ),
    );
  }

  Widget _buildRuntimeCard() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Text('Runtime', style: Theme.of(context).textTheme.titleLarge),
            const SizedBox(height: 12),
            TextField(
              key: const Key('projectRootField'),
              controller: _projectRootController,
              enabled: !_isRunning,
              decoration: const InputDecoration(
                labelText: 'Repository root',
                hintText:
                    r'C:\path\to\Windows-Audio-and-Display-Baseline-Enforcer',
                border: OutlineInputBorder(),
              ),
            ),
            const SizedBox(height: 12),
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Expanded(
                  child: TextField(
                    key: const Key('pythonField'),
                    controller: _pythonController,
                    enabled: !_isRunning,
                    decoration: const InputDecoration(
                      labelText: 'Python command',
                      hintText: 'python',
                      border: OutlineInputBorder(),
                    ),
                  ),
                ),
                const SizedBox(width: 12),
                OutlinedButton.icon(
                  key: const Key('loadConfigButton'),
                  onPressed: _isRunning ? null : _loadConfig,
                  icon: const Icon(Icons.refresh),
                  label: const Text('Load config.py'),
                ),
              ],
            ),
            const SizedBox(height: 16),
            Text(
              'Deployment options',
              style: Theme.of(context).textTheme.titleMedium,
            ),
            const SizedBox(height: 12),
            TextField(
              key: const Key('bgInfoFolderField'),
              controller: _bgInfoFolderController,
              enabled: !_isRunning,
              decoration: const InputDecoration(
                labelText: 'BGInfo folder',
                hintText: '25_26',
                prefixText: r'BGInfo\',
                border: OutlineInputBorder(),
              ),
            ),
            const SizedBox(height: 12),
            TextField(
              key: const Key('workersField'),
              controller: _workersController,
              enabled: !_isRunning,
              keyboardType: TextInputType.number,
              decoration: const InputDecoration(
                labelText: 'Maximum concurrent targets',
                border: OutlineInputBorder(),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildTargetsCard() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Text('Targets', style: Theme.of(context).textTheme.titleLarge),
            const SizedBox(height: 12),
            SegmentedButton<TargetSource>(
              key: const Key('targetSourceSelector'),
              segments: const [
                ButtonSegment(
                  value: TargetSource.file,
                  icon: Icon(Icons.description_outlined),
                  label: Text('targets.txt'),
                ),
                ButtonSegment(
                  value: TargetSource.direct,
                  icon: Icon(Icons.edit_outlined),
                  label: Text('Enter directly'),
                ),
              ],
              selected: {_targetSource},
              onSelectionChanged: _isRunning
                  ? null
                  : (selection) {
                      setState(() => _targetSource = selection.first);
                    },
            ),
            const SizedBox(height: 12),
            if (_targetSource == TargetSource.file)
              Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  TextField(
                    key: const Key('targetsFileEditor'),
                    controller: _targetsFileController,
                    enabled: !_isRunning && !_isUpdatingTargetsFile,
                    minLines: 7,
                    maxLines: 12,
                    decoration: const InputDecoration(
                      labelText: 'targets.txt preview and editor',
                      hintText: 'PC-001\nPC-002\nlocalhost',
                      helperText:
                          'One hostname per line. Changes require Save.',
                      alignLabelWithHint: true,
                      border: OutlineInputBorder(),
                    ),
                  ),
                  const SizedBox(height: 12),
                  Row(
                    mainAxisAlignment: MainAxisAlignment.end,
                    children: [
                      OutlinedButton.icon(
                        key: const Key('reloadTargetsButton'),
                        onPressed: _isRunning || _isUpdatingTargetsFile
                            ? null
                            : _loadTargetsFile,
                        icon: const Icon(Icons.refresh),
                        label: const Text('Reload'),
                      ),
                      const SizedBox(width: 8),
                      FilledButton.icon(
                        key: const Key('saveTargetsButton'),
                        onPressed: _isRunning || _isUpdatingTargetsFile
                            ? null
                            : _saveTargetsFile,
                        icon: const Icon(Icons.save_outlined),
                        label: const Text('Save'),
                      ),
                    ],
                  ),
                ],
              )
            else
              TextField(
                key: const Key('directTargetsField'),
                controller: _targetsController,
                enabled: !_isRunning,
                minLines: 7,
                maxLines: 12,
                decoration: const InputDecoration(
                  labelText: 'Target PCs',
                  hintText: 'PC-001\nPC-002\nlocalhost',
                  helperText: 'Enter one hostname per line.',
                  alignLabelWithHint: true,
                  border: OutlineInputBorder(),
                ),
              ),
          ],
        ),
      ),
    );
  }

  Widget _buildFeaturesCard() {
    final shortcutsAvailable = _audioRecall && _displayRecall;
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Text('Features', style: Theme.of(context).textTheme.titleLarge),
            SwitchListTile(
              key: const Key('audioRecallSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Audio recall'),
              value: _audioRecall,
              onChanged: _isRunning
                  ? null
                  : (value) => setState(() {
                      _audioRecall = value;
                      if (!value || !_displayRecall) {
                        _addDesktopShortcuts = false;
                      }
                    }),
            ),
            SwitchListTile(
              key: const Key('displayRecallSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Display recall'),
              value: _displayRecall,
              onChanged: _isRunning
                  ? null
                  : (value) => setState(() {
                      _displayRecall = value;
                      if (!_audioRecall || !value) {
                        _addDesktopShortcuts = false;
                      }
                    }),
            ),
            SwitchListTile(
              key: const Key('bgInfoSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Install BGInfo'),
              value: _bgInfoInstall,
              onChanged: _isRunning
                  ? null
                  : (value) => setState(() => _bgInfoInstall = value),
            ),
            Tooltip(
              message: shortcutsAvailable
                  ? 'Creates desktop shortcuts for the enabled recall tools.'
                  : 'Enable both Audio recall and Display recall to enable desktop shortcuts.',
              child: SwitchListTile(
                key: const Key('shortcutsSwitch'),
                contentPadding: EdgeInsets.zero,
                title: const Text('Add desktop shortcuts'),
                subtitle: shortcutsAvailable
                    ? null
                    : const Text('Requires Audio recall and Display recall.'),
                value: shortcutsAvailable ? _addDesktopShortcuts : false,
                onChanged: _isRunning || !shortcutsAvailable
                    ? null
                    : (value) => setState(() => _addDesktopShortcuts = value),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Color _pcProgressColor(PcProgress progress) {
    return switch (progress) {
      PcProgress.queued => const Color(0xff94a3b8),
      PcProgress.running => const Color(0xff38bdf8),
      PcProgress.complete => const Color(0xff22c55e),
      PcProgress.offline => const Color(0xff3b82f6),
      PcProgress.warning => const Color(0xffff9800),
      PcProgress.error || PcProgress.fatal => const Color(0xffef4444),
    };
  }

  IconData _pcProgressIcon(PcProgress progress) {
    return switch (progress) {
      PcProgress.queued => Icons.schedule,
      PcProgress.running => Icons.sync,
      PcProgress.complete => Icons.check_circle,
      PcProgress.offline => Icons.cloud_off,
      PcProgress.warning => Icons.warning_amber_rounded,
      PcProgress.error || PcProgress.fatal => Icons.error,
    };
  }

  String _pcProgressLabel(PcProgress progress) {
    return switch (progress) {
      PcProgress.queued => 'Queued',
      PcProgress.running => 'Running',
      PcProgress.complete => 'Complete',
      PcProgress.offline => 'Offline',
      PcProgress.warning => 'Warning',
      PcProgress.error => 'Error',
      PcProgress.fatal => 'Fatal',
    };
  }

  Widget _buildProgressIndicators() {
    if (_knownTargets.isEmpty) {
      return const Padding(
        padding: EdgeInsets.symmetric(vertical: 12),
        child: Text('Target progress will appear when deployment starts.'),
      );
    }
    final progress = _finishedTargets.length / _knownTargets.length;
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        LinearProgressIndicator(value: progress),
        const SizedBox(height: 8),
        Text(
          '${_finishedTargets.length} of ${_knownTargets.length} targets finished',
          key: const Key('overallProgressText'),
        ),
        const SizedBox(height: 12),
        LayoutBuilder(
          builder: (context, constraints) {
            const spacing = 10.0;
            final itemWidth = (constraints.maxWidth - (spacing * 3)) / 4;
            return Wrap(
              spacing: spacing,
              runSpacing: spacing,
              children: [
                for (final pc in _knownTargets)
                  SizedBox(
                    width: itemWidth,
                    child: Builder(
                      builder: (context) {
                        final pcProgress = _pcProgress[pc] ?? PcProgress.queued;
                        final pcColor = _pcProgressColor(pcProgress);
                        return Container(
                          key: Key('pcProgress-$pc'),
                          height: 64,
                          padding: const EdgeInsets.symmetric(horizontal: 8),
                          decoration: BoxDecoration(
                            color: pcColor.withValues(alpha: 0.12),
                            borderRadius: BorderRadius.circular(8),
                            border: Border.all(color: pcColor),
                          ),
                          child: Row(
                            children: [
                              if (pcProgress == PcProgress.running)
                                SizedBox(
                                  width: 20,
                                  height: 20,
                                  child: CircularProgressIndicator(
                                    strokeWidth: 2.5,
                                    color: pcColor,
                                  ),
                                )
                              else
                                Icon(
                                  _pcProgressIcon(pcProgress),
                                  color: pcColor,
                                  size: 20,
                                ),
                              const SizedBox(width: 6),
                              Expanded(
                                child: Tooltip(
                                  message: pc,
                                  child: Column(
                                    mainAxisAlignment: MainAxisAlignment.center,
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      Text(
                                        truncateComputerName(pc),
                                        maxLines: 1,
                                        overflow: TextOverflow.clip,
                                        style: const TextStyle(
                                          fontWeight: FontWeight.w600,
                                        ),
                                      ),
                                      if (pcProgress != PcProgress.running)
                                        Text(
                                          _pcProgressLabel(pcProgress),
                                          style: const TextStyle(fontSize: 12),
                                        ),
                                    ],
                                  ),
                                ),
                              ),
                              if (_finishedTargets.contains(pc) &&
                                  !_isRunning) ...[
                                IconButton(
                                  key: Key('viewLogButton-$pc'),
                                  tooltip: 'View log for $pc',
                                  constraints: const BoxConstraints.tightFor(
                                    width: 32,
                                    height: 32,
                                  ),
                                  padding: EdgeInsets.zero,
                                  onPressed: () =>
                                      _showDetailedLogModal(initialPc: pc),
                                  icon: const Icon(
                                    Icons.article_outlined,
                                    size: 19,
                                  ),
                                ),
                                IconButton(
                                  key: Key('retryPcButton-$pc'),
                                  tooltip: 'Retry $pc',
                                  constraints: const BoxConstraints.tightFor(
                                    width: 32,
                                    height: 32,
                                  ),
                                  padding: EdgeInsets.zero,
                                  onPressed: () => _startPostDeploymentRetry(
                                    target: pc,
                                    refreshDeploymentArea: true,
                                  ),
                                  icon: const Icon(Icons.refresh, size: 19),
                                ),
                              ],
                            ],
                          ),
                        );
                      },
                    ),
                  ),
              ],
            );
          },
        ),
      ],
    );
  }

  Widget _buildRunCard() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            Row(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        'Deployment',
                        style: Theme.of(context).textTheme.titleLarge,
                      ),
                      Text(_status, key: const Key('statusText')),
                    ],
                  ),
                ),
                if (_isRunning)
                  OutlinedButton.icon(
                    key: const Key('stopButton'),
                    onPressed: _stopDeployment,
                    icon: const Icon(Icons.stop),
                    label: const Text('Stop'),
                  )
                else
                  FilledButton.icon(
                    key: const Key('deployButton'),
                    onPressed: _startDeployment,
                    icon: const Icon(Icons.play_arrow),
                    label: const Text('Deploy'),
                  ),
              ],
            ),
            const SizedBox(height: 12),
            _buildProgressIndicators(),
            const SizedBox(height: 12),
            Align(
              alignment: Alignment.centerRight,
              child: OutlinedButton.icon(
                key: const Key('openDetailedLogButton'),
                onPressed: _output.isEmpty && _detailedLogPath == null
                    ? null
                    : _showDetailedLogModal,
                icon: const Icon(Icons.article_outlined),
                label: const Text('Open detailed log'),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
