import 'package:deployment_orchestrator_app/display_configuration.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  test('parses a saved extended-desktop display summary', () {
    final configuration = DisplayConfiguration.fromJson({
      'mode': 'Extended desktop',
      'profile_path': r'C:\ProgramData\CTS\display_config_profile.xml',
      'monitors': [
        {
          'number': 1,
          'name': 'DELL P2314T',
          'width': 1920,
          'height': 1080,
          'x': 0,
          'y': 0,
          'rotation': 0,
          'refresh_rate': 60,
          'primary': true,
        },
        {
          'number': 2,
          'name': 'Extron HDMI',
          'width': 1280,
          'height': 720,
          'x': 1920,
          'y': 278,
          'rotation': 0,
          'refresh_rate': 60.0,
          'primary': false,
        },
      ],
    });

    expect(configuration.mode, 'Extended desktop');
    expect(configuration.monitors, hasLength(2));
    expect(configuration.monitors.first.label, 'DELL P2314T');
    expect(configuration.monitors.first.resolution, '1920 × 1080');
    expect(configuration.monitors.last.x, 1920);
    expect(configuration.monitors.last.y, 278);
  });

  testWidgets('renders arrangement, mode, resolutions, and monitor names', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1200, 1000);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    const monitors = [
      DisplayMonitor(
        number: 1,
        name: 'Main monitor',
        width: 1920,
        height: 1080,
        x: 0,
        y: 0,
        rotation: 0,
        refreshRate: 60,
        primary: true,
      ),
      DisplayMonitor(
        number: 2,
        name: 'Right monitor',
        width: 1080,
        height: 1920,
        x: 1920,
        y: -420,
        rotation: 90,
        refreshRate: 59.94,
        primary: false,
      ),
    ];

    await tester.pumpWidget(
      const MaterialApp(
        home: DisplayConfigurationPage(
          pc: 'PC-001',
          configuration: DisplayConfiguration(
            mode: 'Extended desktop',
            profilePath: r'C:\ProgramData\CTS\display_config_profile.xml',
            monitors: monitors,
          ),
        ),
      ),
    );

    expect(find.text('PC-001 display configuration'), findsOneWidget);
    expect(find.text('Extended desktop'), findsOneWidget);
    expect(find.byKey(const Key('monitorArrangement')), findsOneWidget);
    expect(find.text('Main monitor'), findsOneWidget);
    expect(find.text('Right monitor'), findsOneWidget);
    expect(find.text('1920 × 1080 · 60 Hz'), findsOneWidget);
    expect(find.text('1080 × 1920 · 59.94 Hz'), findsOneWidget);
    expect(find.text('Rotated 90°'), findsOneWidget);
    expect(tester.takeException(), isNull);
  });

  testWidgets('combines duplicated screens into a Windows-style label', (
    tester,
  ) async {
    const monitors = [
      DisplayMonitor(
        number: 1,
        name: 'Display one',
        width: 1920,
        height: 1080,
        x: 0,
        y: 0,
        rotation: 0,
        refreshRate: 60,
        primary: true,
      ),
      DisplayMonitor(
        number: 2,
        name: 'Display two',
        width: 1920,
        height: 1080,
        x: 0,
        y: 0,
        rotation: 0,
        refreshRate: 60,
        primary: true,
      ),
    ];

    await tester.pumpWidget(
      const MaterialApp(
        home: Scaffold(body: MonitorArrangement(monitors: monitors)),
      ),
    );

    expect(find.text('1 | 2'), findsOneWidget);
  });

  testWidgets('renders display details in a closable popup', (tester) async {
    tester.view.physicalSize = const Size(1200, 900);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          body: Builder(
            builder: (context) => FilledButton(
              onPressed: () => showDialog<void>(
                context: context,
                builder: (context) => const DisplayConfigurationPage(
                  pc: 'PC-DIALOG',
                  asDialog: true,
                  configuration: DisplayConfiguration(
                    mode: 'Single display',
                    profilePath: '',
                    monitors: [
                      DisplayMonitor(
                        number: 1,
                        name: 'Lectern display',
                        width: 1920,
                        height: 1080,
                        x: 0,
                        y: 0,
                        rotation: 0,
                        refreshRate: 60,
                        primary: true,
                      ),
                    ],
                  ),
                ),
              ),
              child: const Text('Open'),
            ),
          ),
        ),
      ),
    );

    await tester.tap(find.text('Open'));
    await tester.pumpAndSettle();
    expect(find.byType(AlertDialog), findsOneWidget);
    expect(find.text('Display configuration · PC-DIALOG'), findsOneWidget);
    expect(find.text('Lectern display'), findsOneWidget);

    await tester.tap(find.text('Close'));
    await tester.pumpAndSettle();
    expect(find.byType(AlertDialog), findsNothing);
  });

  testWidgets(
    'reloads details when an older monitoring result has no payload',
    (tester) async {
      var loadCount = 0;
      await tester.pumpWidget(
        MaterialApp(
          home: DisplayConfigurationPage(
            pc: 'PC-STALE',
            configuration: null,
            loader: () async {
              loadCount++;
              return const DisplayConfigurationLoadResult(
                configuration: DisplayConfiguration(
                  mode: 'Only show on display 2',
                  profilePath: r'C:\ProgramData\CTS\display_config_profile.xml',
                  monitors: [
                    DisplayMonitor(
                      number: 2,
                      name: 'Lectern monitor',
                      width: 1920,
                      height: 1080,
                      x: 0,
                      y: 0,
                      rotation: 0,
                      refreshRate: 60,
                      primary: true,
                    ),
                  ],
                ),
              );
            },
          ),
        ),
      );

      expect(find.text('Reading saved display profile…'), findsOneWidget);
      await tester.pumpAndSettle();

      expect(loadCount, 1);
      expect(find.text('Only show on display 2'), findsOneWidget);
      expect(find.text('Lectern monitor'), findsOneWidget);
      expect(
        find.text('The saved display profile could not be read.'),
        findsNothing,
      );
    },
  );
}
