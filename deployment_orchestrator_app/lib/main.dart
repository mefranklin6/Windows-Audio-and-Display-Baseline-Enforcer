import 'dart:async';
import 'dart:io';

import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

import 'app_settings.dart';
import 'native_orchestrator.dart';
import 'update_checker.dart';

void main() {
  runApp(const DeploymentOrchestratorApp());
}

typedef DirectoryPicker = Future<String?> Function(String initialDirectory);
typedef TargetFilePicker = Future<String?> Function(String initialDirectory);
typedef TargetFileLoader = Future<String> Function(String path);
typedef BgInfoAssetValidator = Future<BgInfoFolderValidation> Function(
  Directory folder,
);

Future<String?> _pickDirectory(String initialDirectory) {
  return getDirectoryPath(
    initialDirectory: initialDirectory,
    confirmButtonText: 'Select BGInfo folder',
    canCreateDirectories: false,
  );
}

Future<String?> _pickTargetFile(String initialDirectory) async {
  final file = await openFile(initialDirectory: initialDirectory);
  return file?.path;
}

Future<String> _readTargetFile(String path) => File(path).readAsString();

const targetFileHelp = '''Target file format:

• Plain UTF-8 text
• One computer hostname per line
• Blank lines are allowed
• Lines beginning with # are comments
• Do not separate multiple computers with commas or semicolons

Example:
PC-001
CLASSROOM-02
# Temporarily excluded
localhost''';

class TargetFileValidation {
  const TargetFileValidation({required this.targets, required this.errors});

  final List<String> targets;
  final List<String> errors;

  bool get isValid => errors.isEmpty;
}

TargetFileValidation validateTargetFileContents(String contents) {
  final targets = <String>[];
  final errors = <String>[];
  final seen = <String>{};
  final lines = contents.split(RegExp(r'\r?\n'));
  final validCharacters = RegExp(r'^[A-Za-z0-9_.-]+$');

  for (var index = 0; index < lines.length; index++) {
    final target = lines[index].replaceFirst('\uFEFF', '').trim();
    if (target.isEmpty || target.startsWith('#')) continue;
    final lineNumber = index + 1;
    if (target.contains(',') || target.contains(';')) {
      errors.add(
        'Line $lineNumber contains multiple targets. Put each hostname on its own line.',
      );
      continue;
    }
    if (RegExp(r'\s').hasMatch(target)) {
      errors.add('Line $lineNumber contains whitespace inside the hostname.');
      continue;
    }
    if (!validCharacters.hasMatch(target)) {
      errors.add(
        'Line $lineNumber contains unsupported characters: “$target”. Use letters, numbers, hyphens, periods, or underscores.',
      );
      continue;
    }
    if (target.length > 253) {
      errors.add('Line $lineNumber is longer than 253 characters.');
      continue;
    }
    if (seen.add(target.toLowerCase())) targets.add(target);
  }

  if (targets.isEmpty && errors.isEmpty) {
    errors.add('Add at least one computer hostname.');
  }
  return TargetFileValidation(targets: targets, errors: errors);
}

const bgInfoFolderHelp = '''BGInfo folder: the folder containing your BGInfo assets. It can be located anywhere accessible from this computer.

Place the following in that folder:
• The latest BGInfo64.exe
• One .bgi configuration file
• One compatible image file (.jpg, .jpeg, .png, .bmp, or .gif)

Example: C:\\Deployment Assets\\BGInfo\\25_26''';

class BgInfoFolderValidation {
  const BgInfoFolderValidation(this.errors);

  const BgInfoFolderValidation.valid() : errors = const [];

  final List<String> errors;

  bool get isValid => errors.isEmpty;
}

Future<BgInfoFolderValidation> validateBgInfoFolder(Directory folder) async {
  final executables = <String>[];
  final configurations = <String>[];
  final images = <String>[];
  const imageExtensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'};

  try {
    await for (final entity in folder.list(followLinks: false)) {
      if (entity is! File) continue;
      final name = entity.uri.pathSegments.last;
      final normalizedName = name.toLowerCase();
      if (normalizedName == 'bginfo64.exe') {
        executables.add(name);
      } else if (normalizedName.endsWith('.bgi')) {
        configurations.add(name);
      } else if (imageExtensions.any(normalizedName.endsWith)) {
        images.add(name);
      }
    }
  } on FileSystemException catch (error) {
    return BgInfoFolderValidation([
      'The folder could not be read: ${error.message}',
    ]);
  }

  String? singleFileError(
    List<String> files,
    String description,
    String requiredFile,
  ) {
    if (files.isEmpty) {
      return 'Add one $requiredFile for the $description.';
    }
    if (files.length > 1) {
      return 'Keep only one $requiredFile for the $description; found: ${files.join(', ')}.';
    }
    return null;
  }

  final errors = <String>[
    ?singleFileError(executables, 'BGInfo executable', 'BGInfo64.exe'),
    ?singleFileError(configurations, 'BGInfo configuration', '.bgi file'),
    ?singleFileError(
      images,
      'BGInfo background image',
      'compatible image (.jpg, .jpeg, .png, .bmp, or .gif)',
    ),
  ];
  return BgInfoFolderValidation(errors);
}

ThemeData buildAppTheme(Brightness brightness, {bool highContrast = false}) {
  final isDark = brightness == Brightness.dark;
  final scheme = ColorScheme.fromSeed(
    seedColor: isDark ? const Color(0xff8bb8e8) : const Color(0xff245b93),
    brightness: brightness,
    contrastLevel: highContrast ? 1 : 0.15,
  );
  final base = ThemeData(
    brightness: brightness,
    colorScheme: scheme,
    useMaterial3: true,
  );
  final controlShape = RoundedRectangleBorder(
    borderRadius: BorderRadius.circular(12),
  );
  return base.copyWith(
    scaffoldBackgroundColor: scheme.surfaceContainerLowest,
    focusColor: scheme.primary.withValues(alpha: 0.22),
    hoverColor: scheme.primary.withValues(alpha: 0.08),
    appBarTheme: AppBarTheme(
      elevation: 0,
      scrolledUnderElevation: 1,
      centerTitle: false,
      backgroundColor: scheme.surface,
      foregroundColor: scheme.onSurface,
      surfaceTintColor: scheme.surfaceTint,
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      margin: EdgeInsets.zero,
      color: scheme.surfaceContainerLow,
      surfaceTintColor: Colors.transparent,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(18),
        side: BorderSide(color: scheme.outlineVariant),
      ),
    ),
    dialogTheme: DialogThemeData(
      elevation: 8,
      backgroundColor: scheme.surfaceContainerLow,
      surfaceTintColor: Colors.transparent,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
    ),
    inputDecorationTheme: InputDecorationTheme(
      filled: true,
      fillColor: scheme.surfaceContainerHighest.withValues(alpha: 0.5),
      contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 15),
      border: OutlineInputBorder(
        borderRadius: BorderRadius.circular(12),
        borderSide: BorderSide(color: scheme.outline),
      ),
      enabledBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(12),
        borderSide: BorderSide(color: scheme.outlineVariant),
      ),
      focusedBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(12),
        borderSide: BorderSide(color: scheme.primary, width: 2),
      ),
    ),
    filledButtonTheme: FilledButtonThemeData(
      style: FilledButton.styleFrom(
        minimumSize: const Size(0, 48),
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
        shape: controlShape,
        textStyle: const TextStyle(fontWeight: FontWeight.w600),
      ),
    ),
    outlinedButtonTheme: OutlinedButtonThemeData(
      style: OutlinedButton.styleFrom(
        minimumSize: const Size(0, 48),
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
        shape: controlShape,
        side: BorderSide(color: scheme.outline),
        textStyle: const TextStyle(fontWeight: FontWeight.w600),
      ),
    ),
    textButtonTheme: TextButtonThemeData(
      style: TextButton.styleFrom(
        minimumSize: const Size(0, 44),
        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
        shape: controlShape,
        textStyle: const TextStyle(fontWeight: FontWeight.w600),
      ),
    ),
    iconButtonTheme: IconButtonThemeData(
      style: IconButton.styleFrom(
        minimumSize: const Size.square(44),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      ),
    ),
    dividerTheme: DividerThemeData(color: scheme.outlineVariant, space: 1),
    snackBarTheme: SnackBarThemeData(
      behavior: SnackBarBehavior.floating,
      backgroundColor: scheme.inverseSurface,
      contentTextStyle: TextStyle(color: scheme.onInverseSurface),
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
    ),
    tooltipTheme: TooltipThemeData(
      waitDuration: const Duration(milliseconds: 350),
      showDuration: const Duration(seconds: 5),
      decoration: BoxDecoration(
        color: scheme.inverseSurface,
        borderRadius: BorderRadius.circular(8),
      ),
      textStyle: TextStyle(color: scheme.onInverseSurface),
    ),
  );
}

class DeploymentOrchestratorApp extends StatefulWidget {
  const DeploymentOrchestratorApp({
    this.directoryPicker,
    this.targetFilePicker,
    this.targetFileLoader,
    this.bgInfoAssetValidator,
    this.settingsStore,
    super.key,
  });

  final DirectoryPicker? directoryPicker;
  final TargetFilePicker? targetFilePicker;
  final TargetFileLoader? targetFileLoader;
  final BgInfoAssetValidator? bgInfoAssetValidator;
  final SettingsStore? settingsStore;

  @override
  State<DeploymentOrchestratorApp> createState() =>
      _DeploymentOrchestratorAppState();
}

class _DeploymentOrchestratorAppState extends State<DeploymentOrchestratorApp> {
  bool _darkMode = true;
  bool _settingsLoaded = false;
  Map<String, dynamic>? _initialSettings;
  late final SettingsStore _settingsStore;

  @override
  void initState() {
    super.initState();
    _settingsStore = widget.settingsStore ?? JsonSettingsStore.forCurrentUser();
    _loadSettings();
  }

  Future<void> _loadSettings() async {
    final settings = await _settingsStore.load();
    if (!mounted) return;
    setState(() {
      _initialSettings = settings;
      _darkMode = settings?['dark_mode'] as bool? ?? true;
      _settingsLoaded = true;
    });
  }

  Future<void> _saveSettings(Map<String, dynamic> settings) async {
    await _settingsStore.save(settings);
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Windows Audio and Display Baseline Enforcer Orchestrator',
      theme: buildAppTheme(Brightness.light),
      darkTheme: buildAppTheme(Brightness.dark),
      highContrastTheme: buildAppTheme(Brightness.light, highContrast: true),
      highContrastDarkTheme: buildAppTheme(Brightness.dark, highContrast: true),
      themeMode: _darkMode ? ThemeMode.dark : ThemeMode.light,
      home: _settingsLoaded
          ? DeploymentPage(
              darkMode: _darkMode,
              directoryPicker: widget.directoryPicker ?? _pickDirectory,
              targetFilePicker: widget.targetFilePicker ?? _pickTargetFile,
              targetFileLoader: widget.targetFileLoader ?? _readTargetFile,
              bgInfoAssetValidator:
                  widget.bgInfoAssetValidator ?? validateBgInfoFolder,
              initialSettings: _initialSettings,
              onSettingsChanged: _saveSettings,
              onToggleTheme: () => setState(() => _darkMode = !_darkMode),
            )
          : const Scaffold(body: Center(child: CircularProgressIndicator())),
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
  missingAudioOrDisplayConfiguration,
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
    required this.uninstallRecordedAt,
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
      uninstallRecordedAt: json['uninstall_recorded_at'] as String? ?? '',
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
  final String uninstallRecordedAt;

  bool get isUninstalled => uninstallRecordedAt.isNotEmpty && !ctsDeployed;
}

class UninstallReport {
  const UninstallReport({required this.logFile, required this.pcs});

  factory UninstallReport.fromJson(Map<String, dynamic> json) =>
      UninstallReport(
        logFile: json['log_file'] as String? ?? '',
        pcs: (json['pcs'] as List<dynamic>? ?? const [])
            .map(
              (item) =>
                  UninstallPcResult.fromJson(item as Map<String, dynamic>),
            )
            .toList(),
      );

  final String logFile;
  final List<UninstallPcResult> pcs;
}

class UninstallPcResult {
  const UninstallPcResult({
    required this.pc,
    required this.success,
    required this.error,
  });

  factory UninstallPcResult.fromJson(Map<String, dynamic> json) =>
      UninstallPcResult(
        pc: json['pc'] as String? ?? 'Unknown PC',
        success: json['success'] as bool? ?? false,
        error: json['error'] as String? ?? '',
      );

  final String pc;
  final bool success;
  final String error;
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
    required this.directoryPicker,
    required this.targetFilePicker,
    required this.targetFileLoader,
    required this.bgInfoAssetValidator,
    required this.initialSettings,
    required this.onSettingsChanged,
    required this.onToggleTheme,
    super.key,
  });

  final bool darkMode;
  final DirectoryPicker directoryPicker;
  final TargetFilePicker targetFilePicker;
  final TargetFileLoader targetFileLoader;
  final BgInfoAssetValidator bgInfoAssetValidator;
  final Map<String, dynamic>? initialSettings;
  final Future<void> Function(Map<String, dynamic>) onSettingsChanged;
  final VoidCallback onToggleTheme;

  @override
  State<DeploymentPage> createState() => _DeploymentPageState();
}

class _DeploymentPageState extends State<DeploymentPage> {
  late final TextEditingController _projectRootController;
  late final TextEditingController _targetsFilePathController;
  final TextEditingController _targetsController = TextEditingController();
  final TextEditingController _targetsFileController = TextEditingController();
  final TextEditingController _bgInfoFolderController = TextEditingController();
  final TextEditingController _deploymentSearchController =
      TextEditingController();
  final TextEditingController _monitoringSearchController =
      TextEditingController();
  final TextEditingController _workersController = TextEditingController(
    text: '10',
  );
  final ScrollController _pageScrollController = ScrollController();
  final ScrollController _outputScrollController = ScrollController();

  TargetSource _targetSource = TargetSource.file;
  bool _audioRecall = true;
  bool _displayRecall = true;
  bool _bgInfoInstall = false;
  bool _addDesktopShortcuts = true;
  bool _isRunning = false;
  bool _isMonitoring = false;
  bool _isUninstalling = false;
  bool _isUpdatingTargetsFile = false;
  bool _stopRequested = false;
  bool _monitorStopRequested = false;
  bool _checkingForUpdates = false;
  NativeOrchestrator? _deploymentOrchestrator;
  NativeOrchestrator? _monitoringOrchestrator;
  NativeOrchestrator? _uninstallOrchestrator;
  String _output = '';
  String _status = 'Ready';
  String? _detailedLogPath;
  final Map<String, int> _detailedLogOffsets = {};
  List<String> _knownTargets = const [];
  Map<String, PcProgress> _pcProgress = const {};
  Set<String> _finishedTargets = const {};
  Map<String, List<ScriptDeploymentResult>> _scriptResultsByPc = const {};
  String _monitorStatus = 'Ready';
  String _uninstallStatus = 'Ready';
  List<String> _monitorTargets = const [];
  Set<String> _monitorCompletedTargets = const {};
  List<MonitoringPcResult> _monitorResults = const [];
  MonitoringFilter _monitorFilter = MonitoringFilter.all;
  String _deploymentSearch = '';
  String _monitoringSearch = '';
  Timer? _settingsSaveTimer;
  bool _restoreBgInfoInstall = false;

  bool get _controlsLocked => _isRunning || _isMonitoring || _isUninstalling;

  @override
  void initState() {
    super.initState();
    final settings = widget.initialSettings;
    final savedRoot = settings?['project_root'] as String?;
    final projectRoot =
        savedRoot?.trim().isNotEmpty == true &&
            _isProjectRoot(savedRoot!.trim()) &&
            !_isLegacyInstallPath(savedRoot.trim())
        ? savedRoot.trim()
        : _findProjectRoot();
    _projectRootController = TextEditingController(text: projectRoot);
    final savedTargetsFile = settings?['targets_file'] as String?;
    _targetsFilePathController = TextEditingController(
      text:
          savedTargetsFile?.trim().isNotEmpty == true &&
              !_isLegacyInstallPath(savedTargetsFile!.trim())
          ? savedTargetsFile.trim()
          : _join(projectRoot, 'targets.txt'),
    );
    _targetsController.text = settings?['direct_targets'] as String? ?? '';
    _bgInfoFolderController.text = settings?['bginfo_folder'] as String? ?? '';
    final savedWorkers = settings?['max_workers'];
    if (savedWorkers is int && savedWorkers > 0) {
      _workersController.text = savedWorkers.toString();
    } else if (savedWorkers is String && savedWorkers.trim().isNotEmpty) {
      _workersController.text = savedWorkers;
    }
    _targetSource = settings?['target_source'] == TargetSource.direct.name
        ? TargetSource.direct
        : TargetSource.file;
    _audioRecall = settings?['audio_recall'] as bool? ?? true;
    _displayRecall = settings?['display_recall'] as bool? ?? true;
    _addDesktopShortcuts = settings?['desktop_shortcuts'] as bool? ?? true;
    _restoreBgInfoInstall = settings?['bginfo_install'] as bool? ?? false;
    WidgetsBinding.instance.addPostFrameCallback((_) {
      _loadTargetsFile(silent: true);
      _restoreBgInfoSetting();
    });
  }

  Future<void> _restoreBgInfoSetting() async {
    if (!_restoreBgInfoInstall || !mounted) return;
    final valid = await _validateCurrentBgInfoFolder();
    if (!mounted) return;
    setState(() => _bgInfoInstall = valid);
    if (!valid) _scheduleSettingsSave();
  }

  @override
  void dispose() {
    if (_settingsSaveTimer?.isActive ?? false) {
      _settingsSaveTimer!.cancel();
      unawaited(widget.onSettingsChanged(_settingsPayload()));
    }
    _deploymentOrchestrator?.cancel();
    _monitoringOrchestrator?.cancel();
    _uninstallOrchestrator?.cancel();
    _projectRootController.dispose();
    _targetsFilePathController.dispose();
    _targetsController.dispose();
    _targetsFileController.dispose();
    _bgInfoFolderController.dispose();
    _deploymentSearchController.dispose();
    _monitoringSearchController.dispose();
    _workersController.dispose();
    _pageScrollController.dispose();
    _outputScrollController.dispose();
    super.dispose();
  }

  Map<String, dynamic> _settingsPayload() {
    return {
      'dark_mode': widget.darkMode,
      'project_root': _projectRootController.text.trim(),
      'target_source': _targetSource.name,
      'targets_file': _targetsFilePathController.text.trim(),
      'direct_targets': _targetsController.text,
      'audio_recall': _audioRecall,
      'display_recall': _displayRecall,
      'bginfo_install': _bgInfoInstall,
      'desktop_shortcuts': _addDesktopShortcuts,
      'bginfo_folder': _bgInfoFolderController.text.trim(),
      'max_workers': int.tryParse(_workersController.text.trim()) ?? 10,
    };
  }

  void _scheduleSettingsSave() {
    _settingsSaveTimer?.cancel();
    _settingsSaveTimer = Timer(const Duration(milliseconds: 250), () async {
      try {
        await widget.onSettingsChanged(_settingsPayload());
      } on FileSystemException catch (error) {
        if (mounted) {
          _showMessage('Could not save settings: ${error.message}');
        }
      }
    });
  }

  String _findProjectRoot() {
    final starts = <String>{
      Directory.current.absolute.path,
      File(Platform.resolvedExecutable).parent.absolute.path,
      if (Platform.environment['APPDATA'] case final appData?)
        _join(
          appData,
          'Windows Audio and Display Baseline Enforcer Orchestrator',
        ),
    };

    for (final start in starts) {
      var directory = Directory(start);
      for (var level = 0; level < 8; level++) {
        if (_isProjectRoot(directory.path)) {
          return directory.path;
        }
        final parent = directory.parent;
        if (parent.path == directory.path) break;
        directory = parent;
      }
    }
    return Directory.current.absolute.path;
  }

  bool _isProjectRoot(String path) =>
      File(_join(path, 'utility_scripts\\MonitorTarget.ps1')).existsSync() &&
      Directory(_join(path, 'installer_scripts')).existsSync();

  bool _isLegacyInstallPath(String path) {
    final localAppData = Platform.environment['LOCALAPPDATA'];
    if (localAppData == null || localAppData.isEmpty) return false;
    final legacyRoot = _join(
      _join(localAppData, 'Programs'),
      'Windows Audio and Display Baseline Enforcer',
    );
    final normalizedPath = File(path).absolute.path.toLowerCase();
    final normalizedRoot = Directory(legacyRoot).absolute.path.toLowerCase();
    return normalizedPath == normalizedRoot ||
        normalizedPath.startsWith('$normalizedRoot${Platform.pathSeparator}');
  }

  String _join(String parent, String child) =>
      '$parent${Platform.pathSeparator}$child';

  Future<void> _checkForUpdates() async {
    if (_checkingForUpdates) return;
    setState(() => _checkingForUpdates = true);
    try {
      final update = await checkForUpdates();
      if (!mounted) return;
      await showDialog<void>(
        context: context,
        builder: (dialogContext) => AlertDialog(
          title: Row(
            children: [
              Icon(
                update.updateAvailable
                    ? Icons.system_update_alt
                    : Icons.check_circle_outline,
              ),
              const SizedBox(width: 10),
              Text(
                update.updateAvailable
                    ? 'Update available'
                    : 'You are up to date',
              ),
            ],
          ),
          content: Text(
            update.updateAvailable
                ? 'Version ${update.latestVersion} is available. You have version ${update.currentVersion}.'
                : 'Version ${update.currentVersion} is the latest release.',
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(dialogContext).pop(),
              child: const Text('Close'),
            ),
            if (update.updateAvailable)
              FilledButton.icon(
                onPressed: () async {
                  Navigator.of(dialogContext).pop();
                  try {
                    await openWebUrl(update.preferredUrl);
                  } on Object catch (error) {
                    if (mounted) {
                      _showMessage('Could not open the update: $error');
                    }
                  }
                },
                icon: const Icon(Icons.download_outlined),
                label: Text(
                  update.downloadUrl == null
                      ? 'Open releases'
                      : 'Download installer',
                ),
              ),
          ],
        ),
      );
    } on Object catch (error) {
      if (mounted) _showMessage('Could not check for updates: $error');
    } finally {
      if (mounted) setState(() => _checkingForUpdates = false);
    }
  }

  Future<bool> _selectBgInfoFolder() async {
    final bgInfoRoot = Directory(
      _join(_projectRootController.text.trim(), 'BGInfo'),
    ).absolute;
    final currentValue = _bgInfoFolderController.text.trim();
    final initialDirectory = currentValue.isEmpty
        ? bgInfoRoot.path
        : _bgInfoDirectory(currentValue).absolute.path;
    final selectedPath = await widget.directoryPicker(initialDirectory);
    if (selectedPath == null || selectedPath.trim().isEmpty || !mounted) {
      return false;
    }

    final selected = Directory(selectedPath).absolute;
    final validation = await widget.bgInfoAssetValidator(selected);
    if (!validation.isValid) {
      if (mounted) await _showBgInfoAssetErrorDialog(validation);
      return false;
    }
    setState(() => _bgInfoFolderController.text = selected.path);
    _scheduleSettingsSave();
    return true;
  }

  Directory _bgInfoDirectory(String value) {
    final directory = Directory(value);
    if (directory.isAbsolute) return directory;
    return Directory(
      _join(_join(_projectRootController.text.trim(), 'BGInfo'), value),
    );
  }

  Future<void> _showBgInfoFolderDialog({
    bool enableAfterSelection = false,
  }) async {
    final selected = await showDialog<bool>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.info_outline),
            SizedBox(width: 10),
            Text('BGInfo folder'),
          ],
        ),
        content: const SelectableText(bgInfoFolderHelp),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(dialogContext).pop(false),
            child: const Text('Cancel'),
          ),
          FilledButton.icon(
            key: const Key('bgInfoModalSelectButton'),
            onPressed: () async {
              final didSelect = await _selectBgInfoFolder();
              if (didSelect && dialogContext.mounted) {
                Navigator.of(dialogContext).pop(true);
              }
            },
            icon: const Icon(Icons.folder_open),
            label: const Text('Select folder'),
          ),
        ],
      ),
    );
    if (enableAfterSelection && selected == true && mounted) {
      setState(() => _bgInfoInstall = true);
      _scheduleSettingsSave();
    }
  }

  Future<void> _showBgInfoAssetErrorDialog(BgInfoFolderValidation validation) {
    return showDialog<void>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.error_outline),
            SizedBox(width: 10),
            Text('BGInfo folder needs attention'),
          ],
        ),
        content: SizedBox(
          width: 480,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Text(
                'BGInfo remains disabled until this folder has exactly one of each required asset:',
              ),
              const SizedBox(height: 12),
              SelectableText(
                validation.errors.map((error) => '• $error').join('\n'),
              ),
              const SizedBox(height: 12),
              const Text('Fix the listed files, then select the folder again.'),
            ],
          ),
        ),
        actions: [
          FilledButton(
            key: const Key('bgInfoAssetErrorCloseButton'),
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('OK'),
          ),
        ],
      ),
    );
  }

  Future<bool> _validateCurrentBgInfoFolder() async {
    final folderPath = _bgInfoFolderController.text.trim();
    if (folderPath.isEmpty) return false;
    final folder = _bgInfoDirectory(folderPath);
    final validation = await widget.bgInfoAssetValidator(folder);
    if (validation.isValid) return true;
    if (mounted) await _showBgInfoAssetErrorDialog(validation);
    return false;
  }

  Future<void> _setBgInfoInstall(bool value) async {
    if (!value) {
      setState(() => _bgInfoInstall = false);
      _scheduleSettingsSave();
      return;
    }
    if (_bgInfoFolderController.text.trim().isEmpty) {
      await _showBgInfoFolderDialog(enableAfterSelection: true);
      return;
    }
    if (!await _validateCurrentBgInfoFolder()) return;
    setState(() => _bgInfoInstall = true);
    _scheduleSettingsSave();
  }

  Future<void> _loadTargetsFile({bool silent = false}) async {
    final targetsFile = _selectedTargetsFile;
    setState(() => _isUpdatingTargetsFile = true);
    try {
      final contents = await widget.targetFileLoader(targetsFile.path);
      if (!mounted) return;
      setState(() => _targetsFileController.text = contents);
      final validation = validateTargetFileContents(contents);
      if (!validation.isValid && !silent) {
        await _showTargetFileDialog(
          errors: validation.errors,
          path: targetsFile.path,
        );
      } else if (!silent) {
        _showMessage('Loaded ${targetsFile.path}.');
      }
    } on FileSystemException catch (error) {
      if (!silent) {
        await _showTargetFileDialog(
          errors: ['The file could not be read: ${error.message}'],
          path: targetsFile.path,
        );
      }
    } on FormatException {
      if (!silent) {
        await _showTargetFileDialog(
          errors: const ['Save the file as UTF-8 plain text and try again.'],
          path: targetsFile.path,
        );
      }
    } finally {
      if (mounted) setState(() => _isUpdatingTargetsFile = false);
    }
  }

  Future<void> _saveTargetsFile() async {
    FocusManager.instance.primaryFocus?.unfocus();
    final validation = validateTargetFileContents(_targetsFileController.text);
    if (!validation.isValid) {
      await _showTargetFileDialog(
        errors: validation.errors,
        path: _selectedTargetsFile.path,
      );
      return;
    }
    final targetsFile = _selectedTargetsFile;
    setState(() => _isUpdatingTargetsFile = true);
    try {
      await targetsFile.writeAsString(_targetsFileController.text);
      if (!mounted) return;
      _showMessage('Saved ${targetsFile.path}.');
    } on FileSystemException catch (error) {
      if (!mounted) return;
      _showMessage('Could not save target file: ${error.message}');
    } finally {
      if (mounted) setState(() => _isUpdatingTargetsFile = false);
    }
  }

  File get _selectedTargetsFile {
    final path = _targetsFilePathController.text.trim();
    return File(
      path.isEmpty
          ? _join(_projectRootController.text.trim(), 'targets.txt')
          : path,
    );
  }

  Future<void> _selectTargetsFile() async {
    final currentFile = _selectedTargetsFile.absolute;
    final initialDirectory = currentFile.parent.existsSync()
        ? currentFile.parent.path
        : _projectRootController.text.trim();
    final selectedPath = await widget.targetFilePicker(initialDirectory);
    if (selectedPath == null || selectedPath.trim().isEmpty || !mounted) return;

    final file = File(selectedPath).absolute;
    try {
      final contents = await widget.targetFileLoader(file.path);
      final validation = validateTargetFileContents(contents);
      if (!validation.isValid) {
        await _showTargetFileDialog(errors: validation.errors, path: file.path);
        return;
      }
      if (!mounted) return;
      setState(() {
        _targetsFilePathController.text = file.path;
        _targetsFileController.text = contents;
        _targetSource = TargetSource.file;
      });
      _scheduleSettingsSave();
    } on FileSystemException catch (error) {
      if (mounted) {
        await _showTargetFileDialog(
          errors: ['The file could not be read: ${error.message}'],
          path: file.path,
        );
      }
    } on FormatException {
      if (mounted) {
        await _showTargetFileDialog(
          errors: const ['Save the file as UTF-8 plain text and try again.'],
          path: file.path,
        );
      }
    }
  }

  Future<List<String>?> _loadValidatedTargets() async {
    final file = _selectedTargetsFile;
    try {
      final contents = await widget.targetFileLoader(file.path);
      final validation = validateTargetFileContents(contents);
      if (!validation.isValid) {
        await _showTargetFileDialog(errors: validation.errors, path: file.path);
        return null;
      }
      return validation.targets;
    } on FileSystemException catch (error) {
      await _showTargetFileDialog(
        errors: ['The file could not be read: ${error.message}'],
        path: file.path,
      );
      return null;
    } on FormatException {
      await _showTargetFileDialog(
        errors: const ['Save the file as UTF-8 plain text and try again.'],
        path: file.path,
      );
      return null;
    }
  }

  Future<void> _showTargetFileDialog({
    List<String> errors = const [],
    String? path,
  }) {
    return showDialog<void>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: Row(
          children: [
            Icon(errors.isEmpty ? Icons.help_outline : Icons.error_outline),
            const SizedBox(width: 10),
            Text(
              errors.isEmpty
                  ? 'Target file format'
                  : 'Target file needs attention',
            ),
          ],
        ),
        content: SizedBox(
          width: 520,
          child: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                if (path != null) ...[
                  SelectableText(path),
                  const SizedBox(height: 12),
                ],
                if (errors.isNotEmpty) ...[
                  SelectableText(
                    errors.map((error) => '• $error').join('\n'),
                    key: const Key('targetFileValidationErrors'),
                  ),
                  const SizedBox(height: 16),
                ],
                const SelectableText(targetFileHelp),
              ],
            ),
          ),
        ),
        actions: [
          FilledButton(
            key: const Key('targetFileDialogCloseButton'),
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('OK'),
          ),
        ],
      ),
    );
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

  Future<DeploymentReport?> _startDeployment({
    List<String>? targetsOverride,
  }) async {
    if (_controlsLocked) {
      _showMessage('Wait for the current operation to finish.');
      return null;
    }
    FocusManager.instance.primaryFocus?.unfocus();
    final root = _projectRootController.text.trim();
    final workers = int.tryParse(_workersController.text.trim());
    final directTargets = _directTargets();

    if (!Directory(_join(root, 'installer_scripts')).existsSync()) {
      _showMessage('The installer_scripts folder was not found.');
      return null;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return null;
    }
    if (_bgInfoInstall) {
      if (_bgInfoFolderController.text.trim().isEmpty) {
        await _showBgInfoFolderDialog();
        return null;
      }
      if (!await _validateCurrentBgInfoFolder()) return null;
    }
    if (targetsOverride == null &&
        _targetSource == TargetSource.direct &&
        directTargets.isEmpty) {
      _showMessage('Enter at least one target PC.');
      return null;
    }

    List<String> deploymentTargets;
    if (targetsOverride != null) {
      deploymentTargets = targetsOverride;
    } else if (_targetSource == TargetSource.direct) {
      deploymentTargets = directTargets;
    } else {
      final fileTargets = await _loadValidatedTargets();
      if (fileTargets == null) return null;
      deploymentTargets = fileTargets;
    }
    if (deploymentTargets.isEmpty) {
      _showMessage('No target PCs were provided.');
      return null;
    }
    setState(() {
      _isRunning = true;
      _status = targetsOverride == null
          ? 'Starting deployment…'
          : 'Starting filtered redeployment…';
      _detailedLogPath = null;
      _knownTargets = deploymentTargets;
      _pcProgress = {
        for (final target in deploymentTargets) target: PcProgress.queued,
      };
      _finishedTargets = const {};
      _stopRequested = false;
    });

    final orchestrator = NativeOrchestrator(
      projectRoot: root,
      maxWorkers: workers,
      onLog: _appendOutput,
    );
    _deploymentOrchestrator = orchestrator;
    try {
      setState(() {
        _status = targetsOverride == null
            ? 'Deploying ${deploymentTargets.length} target(s)'
            : 'Redeploying ${deploymentTargets.length} selected PC(s)';
      });
      final report = DeploymentReport.fromJson(
        await orchestrator.deploy(
          deploymentTargets,
          DeploymentOptions(
            audioRecall: _audioRecall,
            displayRecall: _displayRecall,
            bgInfoInstall: _bgInfoInstall,
            desktopShortcuts: _addDesktopShortcuts,
            bgInfoFolder: _bgInfoFolderController.text.trim(),
          ),
        ),
      );
      if (!mounted) return null;
      _appendReportScriptStatuses(report);
      final reportHasProblems = report.pcs.any((result) => result.hasProblems);
      setState(() {
        _deploymentOrchestrator = null;
        _isRunning = false;
        _detailedLogPath = report.logFile;
        _knownTargets = report.pcs.map((result) => result.pc).toList();
        _pcProgress = {
          for (final result in report.pcs)
            result.pc: _progressForResult(result),
        };
        _finishedTargets = report.pcs.map((result) => result.pc).toSet();
        final operationName = targetsOverride == null
            ? 'Deployment'
            : 'Filtered redeployment';
        _status = _stopRequested
            ? '$operationName stopped'
            : reportHasProblems
            ? '$operationName finished with issues'
            : '$operationName finished successfully';
      });
      return report;
    } on Object catch (error) {
      if (!mounted) return null;
      setState(() {
        _deploymentOrchestrator = null;
        _isRunning = false;
        _status = 'Could not start deployment';
      });
      _appendOutput('ERROR: $error');
      return null;
    }
  }

  Future<DeploymentReport?> _startPostDeploymentRetry({
    required String target,
    required bool refreshDeploymentArea,
    ValueChanged<String>? onProgress,
  }) async {
    final root = _projectRootController.text.trim();
    final workers = int.tryParse(_workersController.text.trim());
    if (!Directory(_join(root, 'installer_scripts')).existsSync()) {
      _showMessage('The installer_scripts folder was not found.');
      return null;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return null;
    }
    if (_bgInfoInstall) {
      if (_bgInfoFolderController.text.trim().isEmpty) {
        await _showBgInfoFolderDialog();
        return null;
      }
      if (!await _validateCurrentBgInfoFolder()) return null;
    }

    if (refreshDeploymentArea) {
      setState(() {
        _pcProgress = {..._pcProgress, target: PcProgress.queued};
        _finishedTargets = {..._finishedTargets.where((pc) => pc != target)};
      });
    }

    final orchestrator = NativeOrchestrator(
      projectRoot: root,
      maxWorkers: workers,
      onLog: (line) => _appendOutput(
        line,
        onProgress: onProgress,
        updateProgress: refreshDeploymentArea,
      ),
    );
    try {
      final report = DeploymentReport.fromJson(
        await orchestrator.deploy(
          [target],
          DeploymentOptions(
            audioRecall: _audioRecall,
            displayRecall: _displayRecall,
            bgInfoInstall: _bgInfoInstall,
            desktopShortcuts: _addDesktopShortcuts,
            bgInfoFolder: _bgInfoFolderController.text.trim(),
          ),
        ),
      );
      if (!mounted) return null;
      _appendReportScriptStatuses(report);
      if (!mounted) return report;
      final targetResult = report.pcs.firstWhere(
        (result) => result.pc == target,
        orElse: () => _failedRetryReport(target).pcs.single,
      );
      setState(() {
        if (report.logFile.isNotEmpty) _detailedLogPath = report.logFile;
        if (refreshDeploymentArea) {
          _pcProgress = {
            ..._pcProgress,
            target: _progressForResult(targetResult),
          };
          _finishedTargets = {..._finishedTargets, target};
        }
      });
      return report;
    } on Object catch (error) {
      _appendOutput('ERROR: $error', onProgress: onProgress);
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
    if (_controlsLocked) {
      _showMessage('Wait for the current operation to finish.');
      return;
    }
    FocusManager.instance.primaryFocus?.unfocus();
    final root = _projectRootController.text.trim();
    final workers = int.tryParse(_workersController.text.trim());
    final directTargets = _directTargets();

    if (!File(_join(root, 'utility_scripts\\MonitorTarget.ps1')).existsSync()) {
      _showMessage('utility_scripts\\MonitorTarget.ps1 was not found.');
      return;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return;
    }
    if (_targetSource == TargetSource.direct && directTargets.isEmpty) {
      _showMessage('Enter at least one target PC.');
      return;
    }

    final List<String> targets;
    if (_targetSource == TargetSource.direct) {
      targets = directTargets;
    } else {
      final fileTargets = await _loadValidatedTargets();
      if (fileTargets == null) return;
      targets = fileTargets;
    }
    if (targets.isEmpty) {
      _showMessage('No target PCs were provided.');
      return;
    }
    setState(() {
      _isMonitoring = true;
      _monitorStatus = 'Monitoring ${targets.length} target(s)…';
      _monitorTargets = targets;
      _monitorCompletedTargets = const {};
      _monitorResults = const [];
      _monitorFilter = MonitoringFilter.all;
      _monitorStopRequested = false;
    });

    final orchestrator = NativeOrchestrator(
      projectRoot: root,
      maxWorkers: workers,
      onMonitoringProgress: (pc) {
        if (!mounted) return;
        setState(() {
          _monitorCompletedTargets = {..._monitorCompletedTargets, pc};
        });
      },
    );
    _monitoringOrchestrator = orchestrator;
    try {
      final report = MonitoringReport.fromJson(
        await orchestrator.monitor(targets),
      );
      if (!mounted) return;
      setState(() {
        _monitoringOrchestrator = null;
        _isMonitoring = false;
        _monitorResults = report.pcs;
        _monitorCompletedTargets = _monitorTargets.toSet();
        _monitorStatus = _monitorStopRequested
            ? 'Monitoring stopped'
            : 'Monitoring complete';
      });
    } on Object catch (error) {
      if (!mounted) return;
      setState(() {
        _monitoringOrchestrator = null;
        _isMonitoring = false;
        _monitorStatus = 'Could not start monitoring';
      });
      _showMessage('$error');
    }
  }

  void _stopMonitoring() {
    if (_monitoringOrchestrator != null) {
      _monitoringOrchestrator!.cancel();
      setState(() {
        _monitorStopRequested = true;
        _monitorStatus = 'Stopping monitoring…';
      });
    } else {
      _showMessage('The monitoring process could not be stopped.');
    }
  }

  Future<void> _confirmUninstall() async {
    if (_controlsLocked) {
      _showMessage('Wait for the current operation to finish.');
      return;
    }
    final confirmed = await showDialog<bool>(
      context: context,
      builder: (dialogContext) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.warning_amber_rounded),
            SizedBox(width: 10),
            Text('Uninstall from selected PCs?'),
          ],
        ),
        content: const Text(
          'This removes CTS configuration, PowerShell modules, startup files, '
          'BGInfo files, and CTS desktop shortcuts from every selected PC. '
          'This action cannot be undone automatically.',
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(dialogContext).pop(false),
            child: const Text('Cancel'),
          ),
          FilledButton(
            key: const Key('confirmUninstallButton'),
            onPressed: () => Navigator.of(dialogContext).pop(true),
            style: FilledButton.styleFrom(
              backgroundColor: Theme.of(dialogContext).colorScheme.error,
              foregroundColor: Theme.of(dialogContext).colorScheme.onError,
            ),
            child: const Text('Uninstall'),
          ),
        ],
      ),
    );
    if (confirmed == true && mounted) await _startUninstall();
  }

  Future<void> _startUninstall() async {
    FocusManager.instance.primaryFocus?.unfocus();
    final root = _projectRootController.text.trim();
    final workers = int.tryParse(_workersController.text.trim());
    final directTargets = _directTargets();

    if (!File(_join(root, 'utility_scripts\\uninstall.ps1')).existsSync()) {
      _showMessage('utility_scripts\\uninstall.ps1 was not found.');
      return;
    }
    if (workers == null || workers < 1) {
      _showMessage('Maximum workers must be a whole number of at least 1.');
      return;
    }
    if (_targetSource == TargetSource.direct && directTargets.isEmpty) {
      _showMessage('Enter at least one target PC.');
      return;
    }

    final List<String> targets;
    if (_targetSource == TargetSource.direct) {
      targets = directTargets;
    } else {
      final fileTargets = await _loadValidatedTargets();
      if (fileTargets == null) return;
      targets = fileTargets;
    }
    if (targets.isEmpty) {
      _showMessage('No target PCs were provided.');
      return;
    }

    setState(() {
      _isUninstalling = true;
      _uninstallStatus = 'Uninstalling ${targets.length} target(s)…';
      _detailedLogPath = null;
    });
    final orchestrator = NativeOrchestrator(
      projectRoot: root,
      maxWorkers: workers,
      onLog: (line) => _appendOutput(line, updateProgress: false),
    );
    _uninstallOrchestrator = orchestrator;
    try {
      final report = UninstallReport.fromJson(
        await orchestrator.uninstall(targets),
      );
      if (!mounted) return;
      final failures = report.pcs.where((result) => !result.success).length;
      setState(() {
        _uninstallOrchestrator = null;
        _isUninstalling = false;
        _detailedLogPath = report.logFile;
        _uninstallStatus = failures == 0
            ? 'Uninstall finished successfully'
            : 'Uninstall finished with issues ($failures failed)';
      });
    } on Object catch (error) {
      if (!mounted) return;
      setState(() {
        _uninstallOrchestrator = null;
        _isUninstalling = false;
        _uninstallStatus = 'Could not start uninstall';
      });
      _appendOutput('ERROR: $error', updateProgress: false);
    }
  }

  void _stopUninstall() {
    final orchestrator = _uninstallOrchestrator;
    if (orchestrator == null) {
      _showMessage('The uninstall process could not be stopped.');
      return;
    }
    orchestrator.cancel();
    setState(() => _uninstallStatus = 'Stopping uninstall…');
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
    if (_deploymentOrchestrator != null) {
      _deploymentOrchestrator!.cancel();
      setState(() {
        _stopRequested = true;
        _status = 'Stopping deployment…';
      });
    } else {
      _showMessage('The deployment process could not be stopped.');
    }
  }

  Color _severityColor(String severity) {
    final dark = Theme.of(context).brightness == Brightness.dark;
    return switch (severity.toLowerCase()) {
      'info' => dark ? const Color(0xff6dd58c) : const Color(0xff146c2e),
      'warning' => dark ? const Color(0xffffb95c) : const Color(0xff8a4a00),
      'error' ||
      'fatal' => dark ? const Color(0xffffb4ab) : const Color(0xffb3261e),
      'not_deployed' =>
        dark ? const Color(0xffc4c7c5) : const Color(0xff5f6368),
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

  void _showMessage(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(context)
      ..hideCurrentSnackBar()
      ..showSnackBar(SnackBar(content: Text(message)));
  }

  @override
  Widget build(BuildContext context) {
    final textScale = MediaQuery.textScalerOf(context).scale(1);
    return Scaffold(
      appBar: AppBar(
        toolbarHeight: 72,
        titleSpacing: 20,
        title: Row(
          children: [
            ExcludeSemantics(
              child: Container(
                width: 42,
                height: 42,
                decoration: BoxDecoration(
                  color: Theme.of(context).colorScheme.primaryContainer,
                  borderRadius: BorderRadius.circular(12),
                ),
                child: Icon(
                  Icons.hub_outlined,
                  color: Theme.of(context).colorScheme.onPrimaryContainer,
                ),
              ),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Semantics(
                    header: true,
                    child: Text(
                      'Windows Audio and Display Baseline Enforcer Orchestrator',
                      style: Theme.of(context).textTheme.titleLarge?.copyWith(
                        fontWeight: FontWeight.w700,
                        letterSpacing: -0.2,
                      ),
                    ),
                  ),
                  if (textScale <= 1.5)
                    Text(
                      'Deploy, monitor, and verify CTS endpoints',
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                      style: Theme.of(context).textTheme.bodySmall?.copyWith(
                        color: Theme.of(context).colorScheme.onSurfaceVariant,
                      ),
                    ),
                ],
              ),
            ),
          ],
        ),
        actions: [
          IconButton(
            key: const Key('checkForUpdatesButton'),
            onPressed: _checkingForUpdates ? null : _checkForUpdates,
            tooltip: 'Check for updates (version $applicationVersion)',
            icon: _checkingForUpdates
                ? const SizedBox.square(
                    dimension: 20,
                    child: CircularProgressIndicator(strokeWidth: 2),
                  )
                : const Icon(Icons.system_update_outlined),
          ),
          TextButton.icon(
            key: const Key('settingsButton'),
            onPressed: _controlsLocked ? null : _showSettingsDialog,
            icon: const Icon(Icons.settings_outlined),
            label: const Text('Settings'),
          ),
          IconButton(
            key: const Key('themeToggleButton'),
            onPressed: () {
              widget.onToggleTheme();
              _scheduleSettingsSave();
            },
            tooltip: widget.darkMode
                ? 'Switch to light mode'
                : 'Switch to dark mode',
            icon: Icon(widget.darkMode ? Icons.light_mode : Icons.dark_mode),
          ),
          const SizedBox(width: 8),
        ],
      ),
      body: SafeArea(
        child: FocusTraversalGroup(
          policy: ReadingOrderTraversalPolicy(),
          child: _buildUnifiedView(),
        ),
      ),
    );
  }

  Widget _buildUnifiedView() {
    return Scrollbar(
      controller: _pageScrollController,
      thumbVisibility: true,
      child: SingleChildScrollView(
        controller: _pageScrollController,
        padding: const EdgeInsets.fromLTRB(24, 24, 24, 40),
        child: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 1120),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                LayoutBuilder(
                  builder: (context, constraints) {
                    final textScale = MediaQuery.textScalerOf(context).scale(1);
                    final narrow =
                        constraints.maxWidth < 760 * textScale.clamp(1, 1.5);
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
                const SizedBox(height: 16),
                _buildMonitoringResultsSection(),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _buildSectionHeading({
    required IconData icon,
    required String title,
    String? description,
  }) {
    final colors = Theme.of(context).colorScheme;
    return Row(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        ExcludeSemantics(
          child: Container(
            width: 40,
            height: 40,
            decoration: BoxDecoration(
              color: colors.primaryContainer,
              borderRadius: BorderRadius.circular(11),
            ),
            child: Icon(icon, size: 21, color: colors.onPrimaryContainer),
          ),
        ),
        const SizedBox(width: 12),
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Semantics(
                header: true,
                child: Text(
                  title,
                  style: Theme.of(context).textTheme.titleLarge?.copyWith(
                    fontWeight: FontWeight.w700,
                    letterSpacing: -0.2,
                  ),
                ),
              ),
              if (description != null) ...[
                const SizedBox(height: 2),
                Text(
                  description,
                  style: Theme.of(context).textTheme.bodyMedium
                      ?.copyWith(color: colors.onSurfaceVariant),
                ),
              ],
            ],
          ),
        ),
      ],
    );
  }

  Widget _buildMonitoringResultsSection() {
    final filteredResults = _monitorResults
        .where((result) => _matchesMonitoringFilter(result, _monitorFilter))
        .where((result) => _matchesPcSearch(result.pc, _monitoringSearch))
        .toList();
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            _buildSectionHeading(
              icon: Icons.monitor_heart_outlined,
              title: 'Monitoring results',
              description: 'Read-only checks for CTS files, shortcuts, BGInfo, and installed PowerShell modules.',
            ),
            const SizedBox(height: 12),
            Semantics(
              container: true,
              liveRegion: true,
              label: 'Monitoring status: $_monitorStatus',
              child: ExcludeSemantics(
                child: Text(
                  _monitorStatus,
                  key: const Key('monitoringStatusText'),
                  style: Theme.of(context).textTheme.labelLarge
                      ?.copyWith(color: Theme.of(context).colorScheme.primary),
                ),
              ),
            ),
            if (_monitorResults.isNotEmpty) ...[
              const SizedBox(height: 16),
              const Divider(),
              const SizedBox(height: 8),
              _buildPcSearchField(
                key: const Key('monitoringPcSearchField'),
                controller: _monitoringSearchController,
                label: 'Find a monitored PC',
                choices: _monitorResults.map((result) => result.pc),
                onChanged: (value) => setState(() => _monitoringSearch = value),
              ),
              const SizedBox(height: 10),
              LayoutBuilder(
                builder: (context, constraints) {
                  final dropdown = DropdownButtonFormField<MonitoringFilter>(
                    key: const Key('monitoringReportFilter'),
                    initialValue: _monitorFilter,
                    isExpanded: true,
                    decoration: const InputDecoration(
                      labelText: 'Monitoring report',
                      border: OutlineInputBorder(),
                      isDense: true,
                    ),
                    items: [
                      for (final filter in MonitoringFilter.values)
                        DropdownMenuItem(
                          value: filter,
                          child: Text(_monitoringFilterLabelWithCount(filter)),
                        ),
                    ],
                    onChanged: _controlsLocked
                        ? null
                        : (filter) {
                            if (filter != null) {
                              setState(() => _monitorFilter = filter);
                            }
                          },
                  );
                  final actions = Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      OutlinedButton.icon(
                        key: const Key('copyMonitoringReportButton'),
                        onPressed: filteredResults.isEmpty
                            ? null
                            : () => _copyMonitoringReport(filteredResults),
                        icon: const Icon(Icons.content_copy_rounded, size: 18),
                        label: const Text('Copy PC list'),
                      ),
                      FilledButton.tonalIcon(
                        key: const Key('redeployMonitoringReportButton'),
                        onPressed: _controlsLocked || filteredResults.isEmpty
                            ? null
                            : () => _startDeployment(
                                targetsOverride: filteredResults
                                    .map((result) => result.pc)
                                    .toList(),
                              ),
                        icon: const Icon(Icons.restart_alt_rounded, size: 20),
                        label: const Text('Redeploy selected'),
                      ),
                    ],
                  );
                  final textScale = MediaQuery.textScalerOf(context).scale(1);
                  if (constraints.maxWidth < 760 * textScale.clamp(1, 1.5)) {
                    return Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        dropdown,
                        const SizedBox(height: 10),
                        Align(alignment: Alignment.centerRight, child: actions),
                      ],
                    );
                  }
                  return Row(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      Expanded(child: dropdown),
                      const SizedBox(width: 12),
                      actions,
                    ],
                  );
                },
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
                      child: Text('No PCs match this report and search.'),
                    ),
                  ),
                ),
              LayoutBuilder(
                builder: (context, constraints) {
                  const spacing = 10.0;
                  final itemWidth = (constraints.maxWidth - (spacing * 3)) / 4;
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
              const SizedBox(height: 16),
              Container(
                padding: const EdgeInsets.all(16),
                decoration: BoxDecoration(
                  color: Theme.of(context).colorScheme.surfaceContainerHighest,
                  borderRadius: BorderRadius.circular(10),
                ),
                child: const Text(
                  'Select the targets above, then choose Monitor to inspect them.',
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }

  bool _matchesPcSearch(String pc, String search) =>
      pc.toLowerCase().contains(search.trim().toLowerCase());

  Widget _buildPcSearchField({
    required Key key,
    required TextEditingController controller,
    required String label,
    required Iterable<String> choices,
    required ValueChanged<String> onChanged,
  }) {
    return Autocomplete<String>(
      initialValue: controller.value,
      displayStringForOption: (pc) => pc,
      optionsBuilder: (value) => filterPcChoices(choices, value.text),
      onSelected: (pc) {
        controller.value = TextEditingValue(
          text: pc,
          selection: TextSelection.collapsed(offset: pc.length),
        );
        onChanged(pc);
      },
      fieldViewBuilder: (context, fieldController, focusNode, onSubmitted) {
        return TextField(
          key: key,
          controller: fieldController,
          focusNode: focusNode,
          decoration: InputDecoration(
            labelText: label,
            hintText: 'Start typing a PC name',
            prefixIcon: const Icon(Icons.search),
            suffixIcon: fieldController.text.isEmpty
                ? null
                : IconButton(
                    tooltip: 'Clear PC search',
                    onPressed: () {
                      fieldController.clear();
                      controller.clear();
                      onChanged('');
                      focusNode.requestFocus();
                    },
                    icon: const Icon(Icons.clear),
                  ),
            border: const OutlineInputBorder(),
            isDense: true,
          ),
          onChanged: (value) {
            controller.text = value;
            onChanged(value);
          },
        );
      },
    );
  }

  String _monitoringFilterLabel(MonitoringFilter filter) {
    return switch (filter) {
      MonitoringFilter.all => 'All PCs',
      MonitoringFilter.failedChecks => 'Failed Checks',
      MonitoringFilter.offline => 'Offline PCs',
      MonitoringFilter.winRmUnavailable => 'WinRM Unavailable',
      MonitoringFilter.missingDeployment => 'Missing CTS Deployment',
      MonitoringFilter.missingAudioConfiguration =>
        'Missing Audio Configuration',
      MonitoringFilter.missingDisplayConfiguration =>
        'Missing Display Configuration',
      MonitoringFilter.missingAudioOrDisplayConfiguration =>
        'Missing Audio or Display Configuration',
      MonitoringFilter.missingLogoutShortcut => 'Missing Log Out Shortcut',
      MonitoringFilter.missingRebootShortcut => 'Missing Reboot Shortcut',
      MonitoringFilter.missingBgInfo => 'Missing BGInfo Deployment',
      MonitoringFilter.missingAudioDeviceCmdlets =>
        'Missing AudioDeviceCmdlets',
      MonitoringFilter.missingDisplayConfig => 'Missing DisplayConfig',
    };
  }

  String _monitoringFilterLabelWithCount(MonitoringFilter filter) {
    final count = _monitorResults
        .where((result) => _matchesMonitoringFilter(result, filter))
        .length;
    return '${_monitoringFilterLabel(filter)} ($count)';
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
      MonitoringFilter.missingAudioOrDisplayConfiguration =>
        audioStatus == MonitoringComponentStatus.missing ||
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
    if (result.isUninstalled) return MonitoringComponentStatus.notDeployed;
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
      overallColor = _pcProgressColor(PcProgress.offline);
      overallIcon = Icons.cloud_off;
      overallLabel = 'Offline';
    } else if (!result.winRm) {
      overallColor = _pcProgressColor(PcProgress.error);
      overallIcon = Icons.error;
      overallLabel = 'WinRM unavailable';
    } else if (result.isUninstalled) {
      overallColor = _monitoringStatusColor(
        MonitoringComponentStatus.notDeployed,
      );
      overallIcon = Icons.delete_outline;
      overallLabel = 'Uninstalled';
    } else if (!result.ctsDeployed) {
      overallColor = _pcProgressColor(PcProgress.warning);
      overallIcon = Icons.warning_amber_rounded;
      overallLabel = 'Missing deployment';
    } else {
      final fullyConfigured = !_matchesMonitoringFilter(
        result,
        MonitoringFilter.failedChecks,
      );
      overallColor = _pcProgressColor(
        fullyConfigured ? PcProgress.complete : PcProgress.warning,
      );
      overallIcon = fullyConfigured
          ? Icons.check_circle
          : Icons.warning_amber_rounded;
      overallLabel = fullyConfigured ? 'Healthy' : 'Attention needed';
    }

    return Semantics(
      button: true,
      label: '${result.pc}. Monitoring status: $overallLabel. Open details.',
      child: ExcludeSemantics(
        child: Card(
          key: Key('monitorResult-${result.pc}'),
          clipBehavior: Clip.antiAlias,
          child: InkWell(
            onTap: () => _showMonitoringDetails(result),
            child: ConstrainedBox(
              constraints: const BoxConstraints(minHeight: 84),
              child: Padding(
                padding: const EdgeInsets.symmetric(
                  horizontal: 12,
                  vertical: 12,
                ),
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
                              style: const TextStyle(
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                            Text(
                              overallLabel,
                              maxLines: 1,
                              overflow: TextOverflow.ellipsis,
                              style: Theme.of(context).textTheme.labelMedium
                                  ?.copyWith(color: overallColor),
                            ),
                          ],
                        ),
                      ),
                    ),
                    const Icon(Icons.chevron_right_rounded, size: 20),
                  ],
                ),
              ),
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
                if (result.isUninstalled) ...[
                  Text(
                    'Uninstall recorded: ${result.uninstallRecordedAt}',
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
    final dark = Theme.of(context).brightness == Brightness.dark;
    return switch (status) {
      MonitoringComponentStatus.present =>
        dark ? const Color(0xff6dd58c) : const Color(0xff146c2e),
      MonitoringComponentStatus.missing =>
        dark ? const Color(0xffffb4ab) : const Color(0xffb3261e),
      MonitoringComponentStatus.notDeployed =>
        dark ? const Color(0xffc4c7c5) : const Color(0xff5f6368),
      MonitoringComponentStatus.unknown =>
        dark ? const Color(0xffa8c7fa) : const Color(0xff0b57d0),
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
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            _buildSectionHeading(
              icon: Icons.tune_rounded,
              title: 'Runtime',
              description: 'Application files and deployment concurrency.',
            ),
            const SizedBox(height: 20),
            TextField(
              key: const Key('projectRootField'),
              controller: _projectRootController,
              enabled: !_controlsLocked,
              onChanged: (_) => _scheduleSettingsSave(),
              decoration: const InputDecoration(
                labelText: 'Application files',
                hintText: r'C:\path\to\the installed application',
                helperText: 'The installer configures this automatically. Change it only when running from source.',
                border: OutlineInputBorder(),
              ),
            ),
            const SizedBox(height: 12),
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Expanded(
                  child: TextField(
                    key: const Key('defaultTargetsFileField'),
                    controller: _targetsFilePathController,
                    enabled: !_controlsLocked,
                    readOnly: true,
                    onTap: _controlsLocked ? null : _selectTargetsFile,
                    decoration: InputDecoration(
                      labelText: 'Default target file',
                      hintText: 'Select a plain-text target file',
                      border: const OutlineInputBorder(),
                      suffixIcon: IconButton(
                        key: const Key('settingsTargetsFilePickerButton'),
                        tooltip: 'Select target file',
                        onPressed: _controlsLocked ? null : _selectTargetsFile,
                        icon: const Icon(Icons.file_open_outlined),
                      ),
                    ),
                  ),
                ),
                const SizedBox(width: 8),
                IconButton(
                  key: const Key('targetFileHelpButton'),
                  tooltip: 'Target file format help',
                  onPressed: _showTargetFileDialog,
                  icon: const Icon(Icons.help_outline),
                ),
              ],
            ),
            const SizedBox(height: 16),
            Text(
              'Deployment options',
              style: Theme.of(context).textTheme.titleMedium,
            ),
            const SizedBox(height: 12),
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Expanded(
                  child: TextField(
                    key: const Key('bgInfoFolderField'),
                    controller: _bgInfoFolderController,
                    enabled: !_controlsLocked,
                    readOnly: true,
                    onTap: _controlsLocked ? null : _selectBgInfoFolder,
                    decoration: InputDecoration(
                      labelText: 'BGInfo folder',
                      hintText: 'Select a folder',
                      border: const OutlineInputBorder(),
                      suffixIcon: IconButton(
                        key: const Key('bgInfoFolderPickerButton'),
                        tooltip: 'Select BGInfo folder',
                        onPressed: _controlsLocked ? null : _selectBgInfoFolder,
                        icon: const Icon(Icons.folder_open),
                      ),
                    ),
                  ),
                ),
                const SizedBox(width: 8),
                Tooltip(
                  message: 'BGInfo folder help',
                  child: IconButton(
                    key: const Key('bgInfoHelpButton'),
                    onPressed: _showBgInfoFolderDialog,
                    icon: const Icon(Icons.help_outline),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 12),
            TextField(
              key: const Key('workersField'),
              controller: _workersController,
              enabled: !_controlsLocked,
              keyboardType: TextInputType.number,
              onChanged: (_) => _scheduleSettingsSave(),
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
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            _buildSectionHeading(
              icon: Icons.dns_outlined,
              title: 'Targets',
              description: 'Choose the endpoints for both operations.',
            ),
            const SizedBox(height: 20),
            SegmentedButton<TargetSource>(
              key: const Key('targetSourceSelector'),
              segments: const [
                ButtonSegment(
                  value: TargetSource.file,
                  icon: Icon(Icons.description_outlined),
                  label: Text('From file'),
                ),
                ButtonSegment(
                  value: TargetSource.direct,
                  icon: Icon(Icons.edit_outlined),
                  label: Text('Enter directly'),
                ),
              ],
              selected: {_targetSource},
              onSelectionChanged: _controlsLocked
                  ? null
                  : (selection) {
                      setState(() => _targetSource = selection.first);
                      _scheduleSettingsSave();
                    },
            ),
            const SizedBox(height: 12),
            if (_targetSource == TargetSource.file)
              Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Container(
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: Theme.of(context).colorScheme.surfaceContainerHigh,
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: Row(
                      children: [
                        const Icon(Icons.description_outlined),
                        const SizedBox(width: 10),
                        Expanded(
                          child: Tooltip(
                            message: _selectedTargetsFile.path,
                            child: Text(
                              _selectedTargetsFile.path,
                              key: const Key('selectedTargetsFilePath'),
                              maxLines: 1,
                              overflow: TextOverflow.ellipsis,
                            ),
                          ),
                        ),
                        const SizedBox(width: 8),
                        OutlinedButton.icon(
                          key: const Key('chooseTargetsFileButton'),
                          onPressed: _controlsLocked || _isUpdatingTargetsFile
                              ? null
                              : _selectTargetsFile,
                          icon: const Icon(Icons.file_open_outlined),
                          label: const Text('Choose file'),
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(height: 12),
                  TextField(
                    key: const Key('targetsFileEditor'),
                    controller: _targetsFileController,
                    enabled: !_controlsLocked && !_isUpdatingTargetsFile,
                    minLines: 7,
                    maxLines: 12,
                    decoration: const InputDecoration(
                      labelText: 'Target file preview and editor',
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
                        onPressed: _controlsLocked || _isUpdatingTargetsFile
                            ? null
                            : _loadTargetsFile,
                        icon: const Icon(Icons.refresh),
                        label: const Text('Reload'),
                      ),
                      const SizedBox(width: 8),
                      FilledButton.icon(
                        key: const Key('saveTargetsButton'),
                        onPressed: _controlsLocked || _isUpdatingTargetsFile
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
                enabled: !_controlsLocked,
                onChanged: (_) => _scheduleSettingsSave(),
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
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            _buildSectionHeading(
              icon: Icons.widgets_outlined,
              title: 'Features',
              description: 'Select the CTS components to install or update.',
            ),
            const SizedBox(height: 8),
            SwitchListTile(
              key: const Key('audioRecallSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Audio recall'),
              value: _audioRecall,
              onChanged: _controlsLocked
                  ? null
                  : (value) {
                      setState(() {
                        _audioRecall = value;
                        if (!value || !_displayRecall) {
                          _addDesktopShortcuts = false;
                        }
                      });
                      _scheduleSettingsSave();
                    },
            ),
            SwitchListTile(
              key: const Key('displayRecallSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Display recall'),
              value: _displayRecall,
              onChanged: _controlsLocked
                  ? null
                  : (value) {
                      setState(() {
                        _displayRecall = value;
                        if (!_audioRecall || !value) {
                          _addDesktopShortcuts = false;
                        }
                      });
                      _scheduleSettingsSave();
                    },
            ),
            SwitchListTile(
              key: const Key('bgInfoSwitch'),
              contentPadding: EdgeInsets.zero,
              title: const Text('Install BGInfo'),
              value: _bgInfoInstall,
              onChanged: _controlsLocked ? null : _setBgInfoInstall,
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
                onChanged: _controlsLocked || !shortcutsAvailable
                    ? null
                    : (value) {
                        setState(() => _addDesktopShortcuts = value);
                        _scheduleSettingsSave();
                      },
              ),
            ),
          ],
        ),
      ),
    );
  }

  Color _pcProgressColor(PcProgress progress) {
    final dark = Theme.of(context).brightness == Brightness.dark;
    return switch (progress) {
      PcProgress.queued =>
        dark ? const Color(0xffc4c7c5) : const Color(0xff5f6368),
      PcProgress.running =>
        dark ? const Color(0xff7dd3fc) : const Color(0xff075985),
      PcProgress.complete =>
        dark ? const Color(0xff6dd58c) : const Color(0xff146c2e),
      PcProgress.offline =>
        dark ? const Color(0xffa8c7fa) : const Color(0xff0b57d0),
      PcProgress.warning =>
        dark ? const Color(0xffffb95c) : const Color(0xff8a4a00),
      PcProgress.error || PcProgress.fatal =>
        dark ? const Color(0xffffb4ab) : const Color(0xffb3261e),
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
    final visibleTargets = _knownTargets
        .where((pc) => _matchesPcSearch(pc, _deploymentSearch))
        .toList();
    final progress = _finishedTargets.length / _knownTargets.length;
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        LinearProgressIndicator(
          value: progress,
          minHeight: 8,
          borderRadius: BorderRadius.circular(999),
          semanticsLabel: 'Deployment progress',
          semanticsValue:
              '${_finishedTargets.length} of ${_knownTargets.length} targets finished',
        ),
        const SizedBox(height: 8),
        Text(
          '${_finishedTargets.length} of ${_knownTargets.length} targets finished',
          key: const Key('overallProgressText'),
        ),
        const SizedBox(height: 12),
        _buildPcSearchField(
          key: const Key('deploymentPcSearchField'),
          controller: _deploymentSearchController,
          label: 'Find a deployed PC',
          choices: _knownTargets,
          onChanged: (value) => setState(() => _deploymentSearch = value),
        ),
        const SizedBox(height: 12),
        if (visibleTargets.isEmpty)
          const Card(
            child: Padding(
              padding: EdgeInsets.all(20),
              child: Center(child: Text('No deployed PCs match this search.')),
            ),
          ),
        LayoutBuilder(
          builder: (context, constraints) {
            const spacing = 10.0;
            final itemWidth = (constraints.maxWidth - (spacing * 3)) / 4;
            return Wrap(
              spacing: spacing,
              runSpacing: spacing,
              children: [
                for (final pc in visibleTargets)
                  SizedBox(
                    width: itemWidth,
                    child: Builder(
                      builder: (context) {
                        final pcProgress = _pcProgress[pc] ?? PcProgress.queued;
                        final pcColor = _pcProgressColor(pcProgress);
                        return Container(
                          key: Key('pcProgress-$pc'),
                          constraints: const BoxConstraints(minHeight: 80),
                          padding: const EdgeInsets.symmetric(
                            horizontal: 10,
                            vertical: 8,
                          ),
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
                                    semanticsLabel: '$pc deployment is running',
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
                                child: Semantics(
                                  label:
                                      '$pc. Deployment status: ${_pcProgressLabel(pcProgress)}.',
                                  excludeSemantics: true,
                                  child: Tooltip(
                                    message: pc,
                                    child: Column(
                                      mainAxisAlignment:
                                          MainAxisAlignment.center,
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
                                            style: Theme.of(context)
                                                .textTheme
                                                .labelMedium
                                                ?.copyWith(color: pcColor),
                                          ),
                                      ],
                                    ),
                                  ),
                                ),
                              ),
                              if (_finishedTargets.contains(pc) &&
                                  !_controlsLocked) ...[
                                IconButton(
                                  key: Key('viewLogButton-$pc'),
                                  tooltip: 'View log for $pc',
                                  constraints: const BoxConstraints.tightFor(
                                    width: 40,
                                    height: 40,
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
                                    width: 40,
                                    height: 40,
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
    final monitoringProgress = _monitorTargets.isEmpty
        ? 0.0
        : _monitorCompletedTargets.length / _monitorTargets.length;
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            LayoutBuilder(
              builder: (context, constraints) {
                final summary = Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    _buildSectionHeading(
                      icon: Icons.rocket_launch_outlined,
                      title: 'Operations',
                      description: 'Start one operation at a time using the configuration above.',
                    ),
                    const SizedBox(height: 14),
                    _buildOperationStatus(
                      icon: Icons.rocket_launch_outlined,
                      label: 'Deployment',
                      status: _status,
                      statusKey: const Key('statusText'),
                    ),
                    const SizedBox(height: 4),
                    _buildOperationStatus(
                      icon: Icons.monitor_heart_outlined,
                      label: 'Monitoring',
                      status: _monitorStatus,
                    ),
                    const SizedBox(height: 4),
                    _buildOperationStatus(
                      icon: Icons.delete_outline,
                      label: 'Uninstall',
                      status: _uninstallStatus,
                    ),
                  ],
                );
                final actions = _buildOperationActions();
                final textScale = MediaQuery.textScalerOf(context).scale(1);
                if (constraints.maxWidth < 650 * textScale.clamp(1, 1.5)) {
                  return Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      summary,
                      const SizedBox(height: 14),
                      Align(alignment: Alignment.centerRight, child: actions),
                    ],
                  );
                }
                return Row(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Expanded(child: summary),
                    const SizedBox(width: 16),
                    actions,
                  ],
                );
              },
            ),
            if (_isMonitoring) ...[
              const SizedBox(height: 16),
              LinearProgressIndicator(
                value: monitoringProgress,
                minHeight: 8,
                borderRadius: BorderRadius.circular(999),
                semanticsLabel: 'Monitoring progress',
                semanticsValue:
                    '${_monitorCompletedTargets.length} of ${_monitorTargets.length} targets inspected',
              ),
              const SizedBox(height: 8),
              Text(
                '${_monitorCompletedTargets.length} of ${_monitorTargets.length} targets inspected',
                key: const Key('monitoringProgressText'),
              ),
            ],
            const SizedBox(height: 16),
            const Divider(),
            const SizedBox(height: 12),
            Text(
              'Deployment progress',
              style: Theme.of(context).textTheme.titleMedium,
            ),
            const SizedBox(height: 8),
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

  Widget _buildOperationStatus({
    required IconData icon,
    required String label,
    required String status,
    Key? statusKey,
  }) {
    return Semantics(
      container: true,
      liveRegion: true,
      label: '$label status: $status',
      child: ExcludeSemantics(
        child: Row(
          children: [
            Icon(icon, size: 18, color: Theme.of(context).colorScheme.primary),
            const SizedBox(width: 8),
            Expanded(
              child: Text.rich(
                TextSpan(
                  children: [
                    TextSpan(
                      text: '$label: ',
                      style: const TextStyle(fontWeight: FontWeight.w600),
                    ),
                    TextSpan(text: status),
                  ],
                ),
                key: statusKey,
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildOperationActions() {
    if (_isRunning) {
      return OutlinedButton.icon(
        key: const Key('stopButton'),
        onPressed: _stopDeployment,
        icon: const Icon(Icons.stop),
        label: const Text('Stop deployment'),
      );
    }
    if (_isMonitoring) {
      return OutlinedButton.icon(
        key: const Key('stopMonitoringButton'),
        onPressed: _stopMonitoring,
        icon: const Icon(Icons.stop),
        label: const Text('Stop monitoring'),
      );
    }
    if (_isUninstalling) {
      return OutlinedButton.icon(
        key: const Key('stopUninstallButton'),
        onPressed: _stopUninstall,
        icon: const Icon(Icons.stop),
        label: const Text('Stop uninstall'),
      );
    }
    return ConstrainedBox(
      constraints: const BoxConstraints(maxWidth: 360),
      child: Wrap(
        spacing: 8,
        runSpacing: 8,
        alignment: WrapAlignment.end,
        children: [
          OutlinedButton.icon(
            key: const Key('startMonitoringButton'),
            onPressed: _startMonitoring,
            icon: const Icon(Icons.monitor_heart_rounded, size: 20),
            label: const Text('Monitor'),
          ),
          OutlinedButton.icon(
            key: const Key('uninstallButton'),
            onPressed: _confirmUninstall,
            icon: const Icon(Icons.delete_outline, size: 20),
            label: const Text('Uninstall'),
          ),
          FilledButton.icon(
            key: const Key('deployButton'),
            onPressed: _startDeployment,
            icon: const Icon(Icons.rocket_launch_rounded, size: 20),
            label: const Text('Deploy'),
          ),
        ],
      ),
    );
  }
}
