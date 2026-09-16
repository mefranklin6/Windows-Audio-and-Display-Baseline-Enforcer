import 'package:flutter/material.dart';
import 'package:flutter/semantics.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/main.dart';

void main() {
  testWidgets('shows deployment configuration controls', (tester) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(const DeploymentOrchestratorApp());
    await tester.pump();

    expect(find.text('Deployment Orchestrator'), findsWidgets);
    expect(find.byKey(const Key('audioRecallSwitch')), findsOneWidget);
    expect(find.byKey(const Key('displayRecallSwitch')), findsOneWidget);
    expect(find.byKey(const Key('bgInfoSwitch')), findsOneWidget);
    expect(find.byKey(const Key('shortcutsSwitch')), findsOneWidget);
    expect(find.byKey(const Key('deployButton')), findsOneWidget);
    expect(find.byKey(const Key('startMonitoringButton')), findsOneWidget);
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
    expect(find.byKey(const Key('pythonField')), findsOneWidget);
    expect(find.byKey(const Key('bgInfoFolderField')), findsOneWidget);
    expect(find.byKey(const Key('workersField')), findsOneWidget);
  });

  testWidgets('shows monitoring in the unified operations view', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(const DeploymentOrchestratorApp());
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

    await tester.pumpWidget(const DeploymentOrchestratorApp());
    await tester.pump();

    final monitor = tester.getSemantics(
      find.byKey(const Key('startMonitoringButton')),
    );
    final deploy = tester.getSemantics(find.byKey(const Key('deployButton')));
    final settings = tester.getSemantics(
      find.byKey(const Key('settingsButton')),
    );
    expect(monitor.label, contains('Monitor'));
    expect(monitor.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    expect(deploy.label, contains('Deploy'));
    expect(deploy.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    expect(settings.label, contains('Settings'));
    expect(settings.getSemanticsData().hasAction(SemanticsAction.tap), isTrue);
    semanticsHandle.dispose();
  });

  testWidgets('remains usable at 200 percent text scaling', (tester) async {
    tester.view.physicalSize = const Size(1400, 1100);
    tester.view.devicePixelRatio = 1;
    tester.platformDispatcher.textScaleFactorTestValue = 2;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    addTearDown(tester.platformDispatcher.clearTextScaleFactorTestValue);

    await tester.pumpWidget(const DeploymentOrchestratorApp());
    await tester.pump();

    expect(tester.takeException(), isNull);
    expect(find.byKey(const Key('startMonitoringButton')), findsOneWidget);
    expect(find.byKey(const Key('deployButton')), findsOneWidget);
  });

  testWidgets('toggles between light and dark themes', (tester) async {
    await tester.pumpWidget(const DeploymentOrchestratorApp());
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

    await tester.pumpWidget(const DeploymentOrchestratorApp());
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

    await tester.pumpWidget(const DeploymentOrchestratorApp());
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
  });
}
