import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/semantics.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/app_settings.dart';
import 'package:deployment_orchestrator_app/main.dart';

class MemorySettingsStore implements SettingsStore {
  MemorySettingsStore([Map<String, dynamic>? initial])
    : value = initial == null ? null : Map.of(initial);

  Map<String, dynamic>? value;

  @override
  Future<Map<String, dynamic>?> load() async => value;

  @override
  Future<void> save(Map<String, dynamic> settings) async {
    value = Map.of(settings);
  }
}

DeploymentOrchestratorApp testApp({
  DirectoryPicker? directoryPicker,
  TargetFilePicker? targetFilePicker,
  TargetFileLoader? targetFileLoader,
  BgInfoAssetValidator? bgInfoAssetValidator,
  SettingsStore? settingsStore,
}) {
  return DeploymentOrchestratorApp(
    directoryPicker: directoryPicker,
    targetFilePicker: targetFilePicker,
    targetFileLoader: targetFileLoader ?? (_) async => 'PC-DEFAULT\n',
    bgInfoAssetValidator: bgInfoAssetValidator,
    settingsStore: settingsStore ?? MemorySettingsStore(),
  );
}

void main() {
  testWidgets('shows deployment configuration controls', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pump();

    expect(
      find.text('Windows Audio and Display Baseline Enforcer Orchestrator'),
      findsWidgets,
    );
    expect(find.byKey(const Key('audioRecallSwitch')), findsOneWidget);
    expect(find.byKey(const Key('displayRecallSwitch')), findsOneWidget);
    expect(find.byKey(const Key('bgInfoSwitch')), findsOneWidget);
    expect(find.byKey(const Key('shortcutsSwitch')), findsOneWidget);
    expect(find.byKey(const Key('deployButton')), findsOneWidget);
    expect(find.byKey(const Key('startMonitoringButton')), findsOneWidget);
    expect(find.byKey(const Key('uninstallButton')), findsOneWidget);
    expect(find.byIcon(Icons.monitor_heart_rounded), findsOneWidget);
    expect(find.byIcon(Icons.rocket_launch_rounded), findsOneWidget);
    expect(find.byType(TabBar), findsNothing);
    expect(find.byKey(const Key('projectRootField')), findsNothing);
    expect(find.byKey(const Key('bgInfoFolderField')), findsNothing);
    expect(find.byKey(const Key('workersField')), findsNothing);
    expect(find.byKey(const Key('outputText')), findsNothing);

    await tester.tap(find.byKey(const Key('settingsButton')));
    await tester.pumpAndSettle();
    expect(find.byKey(const Key('projectRootField')), findsOneWidget);
    expect(find.byKey(const Key('pythonField')), findsNothing);
    expect(find.byKey(const Key('bgInfoFolderField')), findsOneWidget);
    expect(find.byKey(const Key('bgInfoFolderPickerButton')), findsOneWidget);
    expect(find.byKey(const Key('bgInfoHelpButton')), findsOneWidget);
    expect(find.byKey(const Key('defaultTargetsFileField')), findsOneWidget);
    expect(
      find.byKey(const Key('settingsTargetsFilePickerButton')),
      findsOneWidget,
    );
    expect(find.byKey(const Key('targetFileHelpButton')), findsOneWidget);
    expect(find.byKey(const Key('workersField')), findsOneWidget);
  });

  testWidgets('uses built-in feature defaults and an empty BGInfo folder', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pump();

    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('audioRecallSwitch')))
          .value,
      isTrue,
    );
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('displayRecallSwitch')))
          .value,
      isTrue,
    );
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('bgInfoSwitch')))
          .value,
      isFalse,
    );
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('shortcutsSwitch')))
          .value,
      isTrue,
    );

    await tester.tap(find.byKey(const Key('settingsButton')));
    await tester.pumpAndSettle();
    final folderField = tester.widget<TextField>(
      find.byKey(const Key('bgInfoFolderField')),
    );
    expect(folderField.controller?.text, isEmpty);
    expect(folderField.readOnly, isTrue);
  });

  testWidgets('prompts for a BGInfo folder before enabling the feature', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    var pickerCalls = 0;

    await tester.pumpWidget(
      testApp(
        directoryPicker: (initialDirectory) async {
          pickerCalls++;
          return '$initialDirectory${Platform.pathSeparator}Classroom';
        },
        bgInfoAssetValidator: (_) async => const BgInfoFolderValidation.valid(),
      ),
    );
    await tester.pump();

    await tester.tap(find.byKey(const Key('bgInfoSwitch')));
    await tester.pumpAndSettle();
    expect(find.text('BGInfo folder'), findsOneWidget);
    expect(find.text(bgInfoFolderHelp), findsOneWidget);
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('bgInfoSwitch')))
          .value,
      isFalse,
    );

    await tester.tap(find.byKey(const Key('bgInfoModalSelectButton')));
    await tester.pumpAndSettle();
    expect(pickerCalls, 1);
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('bgInfoSwitch')))
          .value,
      isTrue,
    );

    await tester.tap(find.byKey(const Key('settingsButton')));
    await tester.pumpAndSettle();
    expect(
      tester
          .widget<TextField>(find.byKey(const Key('bgInfoFolderField')))
          .controller
          ?.text,
      'Classroom',
    );
    await tester.tap(find.byKey(const Key('bgInfoHelpButton')));
    await tester.pumpAndSettle();
    expect(find.text(bgInfoFolderHelp), findsOneWidget);
  });

  testWidgets('selects the BGInfo folder from Settings', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(
      testApp(
        directoryPicker: (initialDirectory) async =>
            '$initialDirectory${Platform.pathSeparator}Auditorium',
        bgInfoAssetValidator: (_) async => const BgInfoFolderValidation.valid(),
      ),
    );
    await tester.pump();
    await tester.tap(find.byKey(const Key('settingsButton')));
    await tester.pumpAndSettle();
    await tester.tap(find.byKey(const Key('bgInfoFolderPickerButton')));
    await tester.pumpAndSettle();

    expect(
      tester
          .widget<TextField>(find.byKey(const Key('bgInfoFolderField')))
          .controller
          ?.text,
      'Auditorium',
    );
  });

  testWidgets('keeps BGInfo disabled and explains invalid assets', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(
      testApp(
        directoryPicker: (initialDirectory) async =>
            '$initialDirectory${Platform.pathSeparator}Incomplete',
        bgInfoAssetValidator: (_) async => const BgInfoFolderValidation([
          'Add one BGInfo64.exe for the BGInfo executable.',
          'Add one .bgi file for the BGInfo configuration.',
          'Add one compatible image (.jpg, .jpeg, .png, .bmp, or .gif) for the BGInfo background image.',
        ]),
      ),
    );
    await tester.pump();

    await tester.tap(find.byKey(const Key('bgInfoSwitch')));
    await tester.pumpAndSettle();
    await tester.tap(find.byKey(const Key('bgInfoModalSelectButton')));
    await tester.pumpAndSettle();

    expect(find.text('BGInfo folder needs attention'), findsOneWidget);
    expect(find.textContaining('Add one BGInfo64.exe'), findsOneWidget);
    expect(find.textContaining('Add one compatible image'), findsOneWidget);
    expect(
      tester
          .widget<SwitchListTile>(find.byKey(const Key('bgInfoSwitch')))
          .value,
      isFalse,
    );
  });

  test('validates required BGInfo assets', () async {
    final folder = await Directory.systemTemp.createTemp('bginfo-validation-');
    addTearDown(() => folder.delete(recursive: true));

    final empty = await validateBgInfoFolder(folder);
    expect(empty.isValid, isFalse);
    expect(empty.errors, hasLength(3));

    await File('${folder.path}${Platform.pathSeparator}BGInfo64.exe')
        .writeAsString('');
    await File('${folder.path}${Platform.pathSeparator}profile.bgi')
        .writeAsString('');
    await File('${folder.path}${Platform.pathSeparator}background.png')
        .writeAsString('');
    expect((await validateBgInfoFolder(folder)).isValid, isTrue);

    await File('${folder.path}${Platform.pathSeparator}second.bgi')
        .writeAsString('');
    final duplicate = await validateBgInfoFolder(folder);
    expect(duplicate.isValid, isFalse);
    expect(duplicate.errors.single, contains('Keep only one .bgi file'));
  });

  test('validates target files as one hostname per line', () {
    final valid = validateTargetFileContents(
      'PC-001\n# Maintenance group\nlocalhost\nPC-001\n',
    );
    expect(valid.isValid, isTrue);
    expect(valid.targets, ['PC-001', 'localhost']);

    final invalid = validateTargetFileContents('PC-001,PC-002\nbad target');
    expect(invalid.isValid, isFalse);
    expect(invalid.errors.first, contains('each hostname on its own line'));
    expect(invalid.errors.last, contains('whitespace'));
  });

  testWidgets('selects and persists a valid target file', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final targetFilePath =
        '${Directory.current.path}${Platform.pathSeparator}classrooms.list';
    final settings = MemorySettingsStore();

    await tester.pumpWidget(
      testApp(
        targetFilePicker: (_) async => targetFilePath,
        targetFileLoader: (_) async => 'PC-001\nPC-002\n',
        settingsStore: settings,
      ),
    );
    await tester.pump();
    await tester.tap(find.byKey(const Key('chooseTargetsFileButton')));
    await tester.runAsync(
      () => Future<void>.delayed(const Duration(milliseconds: 50)),
    );
    await tester.pump();
    await tester.pump(const Duration(milliseconds: 300));

    expect(find.text(File(targetFilePath).absolute.path), findsOneWidget);
    expect(
      tester
          .widget<TextField>(find.byKey(const Key('targetsFileEditor')))
          .controller
          ?.text,
      'PC-001\nPC-002\n',
    );
    expect(settings.value?['targets_file'], File(targetFilePath).absolute.path);
    expect(settings.value?['target_source'], 'file');
  });

  testWidgets('rejects an invalid selected target file with guidance', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final targetFilePath =
        '${Directory.current.path}${Platform.pathSeparator}bad.csv';

    await tester.pumpWidget(
      testApp(
        targetFilePicker: (_) async => targetFilePath,
        targetFileLoader: (_) async => 'PC-001,PC-002',
      ),
    );
    await tester.pump();
    await tester.tap(find.byKey(const Key('chooseTargetsFileButton')));
    await tester.runAsync(
      () => Future<void>.delayed(const Duration(milliseconds: 50)),
    );
    await tester.pump();

    expect(find.text('Target file needs attention'), findsOneWidget);
    expect(
      find.textContaining('each hostname on its own line'),
      findsOneWidget,
    );
    expect(find.text(File(targetFilePath).absolute.path), findsOneWidget);
    expect(find.text(targetFileHelp), findsOneWidget);
  });

  testWidgets('restores and updates persisted settings', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final settings = MemorySettingsStore({
      'dark_mode': false,
      'project_root': r'C:\CTS',
      'target_source': 'direct',
      'targets_file': r'C:\Lists\rooms.txt',
      'direct_targets': 'PC-SAVED',
      'audio_recall': true,
      'display_recall': true,
      'bginfo_install': false,
      'desktop_shortcuts': true,
      'bginfo_folder': 'SavedAssets',
      'max_workers': 4,
    });

    await tester.pumpWidget(testApp(settingsStore: settings));
    await tester.pumpAndSettle();
    expect(
      Theme.of(tester.element(find.byType(Scaffold))).brightness,
      Brightness.light,
    );
    expect(find.text('PC-SAVED'), findsOneWidget);

    await tester.enterText(
      find.byKey(const Key('directTargetsField')),
      'PC-UPDATED',
    );
    await tester.tap(find.byKey(const Key('audioRecallSwitch')));
    await tester.pump(const Duration(milliseconds: 300));
    expect(settings.value?['direct_targets'], 'PC-UPDATED');
    expect(settings.value?['audio_recall'], isFalse);
    expect(settings.value?['targets_file'], r'C:\Lists\rooms.txt');

    await tester.tap(find.byKey(const Key('settingsButton')));
    await tester.pumpAndSettle();
    expect(
      tester
          .widget<TextField>(find.byKey(const Key('workersField')))
          .controller
          ?.text,
      '4',
    );
    expect(
      tester
          .widget<TextField>(find.byKey(const Key('bgInfoFolderField')))
          .controller
          ?.text,
      'SavedAssets',
    );
  });

  test('writes settings to disk as JSON', () async {
    final folder = await Directory.systemTemp.createTemp('settings-store-');
    addTearDown(() => folder.delete(recursive: true));
    final store = JsonSettingsStore(
      File('${folder.path}${Platform.pathSeparator}settings.json'),
    );
    final expected = <String, dynamic>{
      'dark_mode': false,
      'max_workers': 7,
      'targets_file': r'C:\Lists\targets.txt',
    };

    await store.save(expected);

    expect(await store.load(), expected);
    expect(store.file.existsSync(), isTrue);
  });

  testWidgets('shows monitoring in the unified operations view', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pump();

    expect(find.text('Operations'), findsOneWidget);
    expect(find.text('Monitoring results'), findsOneWidget);
    expect(
      find.text(
        'Select the targets above, then choose Monitor to inspect them.',
      ),
      findsOneWidget,
    );
    expect(find.byKey(const Key('startMonitoringButton')), findsOneWidget);
  });

  testWidgets('exposes primary actions to assistive technologies', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final semanticsHandle = tester.ensureSemantics();

    await tester.pumpWidget(testApp());
    await tester.pump();

    final monitor = tester.getSemantics(
      find.byKey(const Key('startMonitoringButton')),
    );
    final deploy = tester.getSemantics(find.byKey(const Key('deployButton')));
    final uninstall = tester.getSemantics(
      find.byKey(const Key('uninstallButton')),
    );
    final settings = tester.getSemantics(
      find.byKey(const Key('settingsButton')),
    );
    expect(monitor.label, contains('Monitor'));
    expect(monitor.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    expect(deploy.label, contains('Deploy'));
    expect(deploy.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    expect(uninstall.label, contains('Uninstall'));
    expect(uninstall.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    expect(settings.label, contains('Settings'));
    expect(settings.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    semanticsHandle.dispose();
  });

  testWidgets('warns before uninstalling selected PCs', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pump();
    await tester.tap(find.byKey(const Key('uninstallButton')));
    await tester.pumpAndSettle();

    expect(find.text('Uninstall from selected PCs?'), findsOneWidget);
    expect(find.byKey(const Key('confirmUninstallButton')), findsOneWidget);
    expect(find.text('Cancel'), findsOneWidget);
  });

  testWidgets('remains usable at 200 percent text scaling', (tester) async {
    tester.view.physicalSize = const Size(1400, 1100);
    tester.view.devicePixelRatio = 1;
    tester.platformDispatcher.textScaleFactorTestValue = 2;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    addTearDown(tester.platformDispatcher.clearTextScaleFactorTestValue);

    await tester.pumpWidget(testApp());
    await tester.pump();

    expect(tester.takeException(), isNull);
    expect(find.byKey(const Key('startMonitoringButton')), findsOneWidget);
    expect(find.byKey(const Key('deployButton')), findsOneWidget);
  });

  testWidgets('toggles between light and dark themes', (tester) async {
    await tester.pumpWidget(testApp());
    await tester.pump();

    expect(
      Theme.of(tester.element(find.byType(Scaffold))).brightness,
      Brightness.dark,
    );
    await tester.tap(find.byKey(const Key('themeToggleButton')));
    await tester.pumpAndSettle();
    expect(
      Theme.of(tester.element(find.byType(Scaffold))).brightness,
      Brightness.light,
    );
  });

  testWidgets('previews targets.txt with reload and save controls', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pumpAndSettle();

    expect(find.byKey(const Key('targetsFileEditor')), findsOneWidget);
    expect(find.byKey(const Key('reloadTargetsButton')), findsOneWidget);
    expect(find.byKey(const Key('saveTargetsButton')), findsOneWidget);
  });

  testWidgets('allows targets to be entered without targets.txt', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(testApp());
    await tester.pump();
    await tester.tap(find.text('Enter directly'));
    await tester.pump();

    final targetField = find.byKey(const Key('directTargetsField'));
    expect(targetField, findsOneWidget);
    await tester.enterText(targetField, 'PC-001\nPC-002');
    expect(find.text('PC-001\nPC-002'), findsOneWidget);
  });

  test('colors severity words and makes fatal bold', () {
    final spans = buildSeveritySpans(
      'INFO ok WARNING careful ERROR failed FATAL stopped',
    );
    final info = spans.firstWhere((span) => span.text == 'INFO');
    final warning = spans.firstWhere((span) => span.text == 'WARNING');
    final error = spans.firstWhere((span) => span.text == 'ERROR');
    final fatal = spans.firstWhere((span) => span.text == 'FATAL');

    expect(info.style?.color, const Color(0xff22c55e));
    expect(warning.style?.color, const Color(0xffff9800));
    expect(error.style?.color, const Color(0xffef4444));
    expect(fatal.style?.color, const Color(0xffef4444));
    expect(fatal.style?.fontWeight, FontWeight.bold);
  });

  test('filters detailed logs by PC and finds matching PC choices', () {
    const log = '''2026 INFO PC-001: Install complete
2026 ERROR PC-002: Display driver failed
2026 WARNING PC-001: Restart recommended''';

    expect(
      filterDetailedLog(log, selectedPc: 'PC-001'),
      '''2026 INFO PC-001: Install complete
2026 WARNING PC-001: Restart recommended''',
    );
    expect(filterPcChoices(['PC-001', 'PC-002', 'LAB-003'], 'pc-002'), [
      'PC-002',
    ]);
  });

  test('fills in non-deployed script statuses', () {
    final statuses = completeScriptStatusList(const [
      ScriptDeploymentResult(
        name: r'installer_scripts\Cleanup.ps1',
        severity: 'info',
        messages: [],
      ),
    ]);

    expect(statuses, hasLength(allDeploymentScriptNames.length));
    expect(
      statuses
          .firstWhere((script) => scriptFileName(script.name) == 'Cleanup.ps1')
          .severity,
      'info',
    );
    expect(
      statuses
          .firstWhere(
            (script) => script.name == 'InstallAudioDeviceCmdlets.ps1',
          )
          .severity,
      'not_deployed',
    );
  });

  test('truncates displayed computer names to fifteen characters', () {
    expect(truncateComputerName('PC-123'), 'PC-123');
    expect(
      truncateComputerName('COMPUTER-NAME-THAT-IS-LONG'),
      'COMPUTER-NAM...',
    );
    expect(truncateComputerName('COMPUTER-NAM...'), hasLength(15));
  });

  test('parses grouped per-PC deployment results', () {
    final report = DeploymentReport.fromJson({
      'log_file': r'C:\logs\deployment.log',
      'pcs': [
        {
          'pc': 'PC-001',
          'highest_severity': 'warning',
          'issues': <Map<String, String>>[],
          'scripts': [
            {
              'name': r'installer_scripts\Cleanup.ps1',
              'severity': 'warning',
              'messages': ['Restart recommended'],
            },
          ],
        },
      ],
    });

    expect(report.logFile, r'C:\logs\deployment.log');
    expect(report.pcs.single.pc, 'PC-001');
    expect(report.pcs.single.hasProblems, isTrue);
    expect(report.pcs.single.scripts.single.severity, 'warning');
    expect(report.pcs.single.scripts.single.messages, ['Restart recommended']);
  });

  test('parses monitoring results and BGInfo startup method', () {
    final report = MonitoringReport.fromJson({
      'log_file': r'C:\logs\monitor.log',
      'pcs': [
        {
          'pc': 'PC-001',
          'online': true,
          'winrm': true,
          'cts_deployed': true,
          'audio_configured': false,
          'display_configured': true,
          'logout_shortcut': true,
          'reboot_shortcut': true,
          'bginfo_deployed': true,
          'bginfo_executable': true,
          'bginfo_profile': true,
          'bginfo_background': true,
          'bginfo_startup': true,
          'bginfo_startup_method': 'av_config_recall.bat',
          'audio_device_cmdlets_versions': ['3.2', '3.3'],
          'display_config_versions': ['6.0.1'],
          'deployment_intent': {
            'recorded_at': '2026-09-14T10:00:00-07:00',
            'audio_recall': false,
            'display_recall': true,
            'bginfo_install': false,
            'desktop_shortcuts': false,
          },
        },
      ],
    });

    expect(report.pcs.single.bgInfoStartupMethod, 'av_config_recall.bat');
    expect(report.pcs.single.audioConfigured, isFalse);
    expect(report.pcs.single.audioDeviceCmdletsVersions, ['3.2', '3.3']);
    expect(report.pcs.single.deploymentIntent?.audioRecall, isFalse);
    expect(report.pcs.single.deploymentIntent?.displayRecall, isTrue);
    expect(report.pcs.single.isUninstalled, isFalse);
  });

  test('recognizes an uninstall record when CTS is absent', () {
    final report = MonitoringReport.fromJson({
      'pcs': [
        {
          'pc': 'PC-REMOVED',
          'online': true,
          'winrm': true,
          'cts_deployed': false,
          'uninstall_recorded_at': '2026-09-17T12:00:00-07:00',
        },
      ],
    });

    expect(report.pcs.single.isUninstalled, isTrue);
  });
}
