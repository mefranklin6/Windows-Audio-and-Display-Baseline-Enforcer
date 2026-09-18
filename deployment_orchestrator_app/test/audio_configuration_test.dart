import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/audio_configuration.dart';

const _levelsJson = '''{
  "PlaybackCommunicationMute": false,
  "PlaybackCommunicationVolume": "55.04457%",
  "PlaybackMute": false,
  "PlaybackVolume": "55.04457%",
  "RecordingCommunicationMute": false,
  "RecordingCommunicationVolume": "80.99999%",
  "RecordingMute": false,
  "RecordingVolume": "80.99999%"
}''';

const _legacyLevelsJson = '''{
  "PlaybackCommunicationVolume": "55.04457%",
  "PlaybackVolume": "55.04457%",
  "RecordingCommunicationVolume": "80.99999%",
  "RecordingVolume": "80.99999%"
}''';

const _devicesJson = '''[
  {
    "Index": 1,
    "Enabled": true,
    "Default": true,
    "DefaultCommunication": true,
    "Type": "Playback",
    "Name": "Headphones",
    "ID": "playback-1",
    "UnrecognizedField": "preserved"
  },
  {
    "Index": 2,
    "Enabled": true,
    "Default": false,
    "DefaultCommunication": false,
    "Type": "Playback",
    "Name": "HDMI",
    "ID": "playback-2"
  },
  {
    "Index": 3,
    "Enabled": true,
    "Default": true,
    "DefaultCommunication": true,
    "Type": "Recording",
    "Name": "Microphone",
    "ID": "recording-1"
  }
]''';

class MemoryAudioGateway implements AudioConfigurationGateway {
  String levelsJson = _levelsJson;
  String devicesJson = _devicesJson;
  String? log = '2026-09-18 [INFO] Audio restored';
  int saves = 0;

  @override
  Future<AudioConfigurationData> load(String pc) async {
    return AudioConfigurationData.fromJsonText(
      levelsJson: levelsJson,
      devicesJson: devicesJson,
      log: log,
    );
  }

  @override
  Future<void> save(
    String pc, {
    required String levelsJson,
    required String devicesJson,
  }) async {
    AudioConfigurationData.fromJsonText(
      levelsJson: levelsJson,
      devicesJson: devicesJson,
    );
    this.levelsJson = levelsJson;
    this.devicesJson = devicesJson;
    saves++;
  }
}

void main() {
  test('decodes PowerShell UTF-16 JSON and preserves its encoding', () {
    const text = '{"Name":"Microphone – Room A"}';

    for (final encoding in AudioTextEncoding.values) {
      final encoded = encodeAudioText(text, encoding);
      final decoded = decodeAudioText(encoded);

      expect(decoded.text, text);
      expect(decoded.encoding, encoding);
    }
  });

  test('parses percent volumes and preserves full device records', () {
    final data = AudioConfigurationData.fromJsonText(
      levelsJson: _levelsJson,
      devicesJson: _devicesJson,
    );

    expect(
      parseAudioVolume(data.levels['RecordingVolume'], key: 'RecordingVolume'),
      closeTo(80.99999, 0.00001),
    );
    expect(data.devices.first['Name'], 'Headphones');
    expect(data.devices.first['Type'], 'Playback');
    expect(data.devices.first['UnrecognizedField'], 'preserved');
  });

  test('rejects invalid raw volume values', () {
    expect(
      () => AudioConfigurationData.fromJsonText(
        levelsJson: _levelsJson.replaceFirst('55.04457%', '101%'),
        devicesJson: _devicesJson,
      ),
      throwsFormatException,
    );
  });

  testWidgets('shows unrecorded legacy mute states as unknown', (tester) async {
    tester.view.physicalSize = const Size(1100, 900);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final gateway = MemoryAudioGateway()..levelsJson = _legacyLevelsJson;

    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          body: AudioConfigurationDialog(pc: 'PC-LEGACY', gateway: gateway),
        ),
      ),
    );
    await tester.pumpAndSettle();

    expect(find.text('Audio configuration · PC-LEGACY'), findsOneWidget);
    expect(find.text('Unknown'), findsNWidgets(4));
    expect(find.byIcon(Icons.help_outline_rounded), findsNWidgets(4));
    expect(find.textContaining('Could not load'), findsNothing);

    await tester.tap(find.byKey(const Key('editAudioButton')));
    await tester.pump();
    await tester.tap(find.byKey(const Key('saveAudioButton')));
    await tester.pumpAndSettle();
    expect(gateway.levelsJson, isNot(contains('Mute')));
  });

  testWidgets('edits GUI values, exposes JSON, and opens the audio log', (
    tester,
  ) async {
    tester.view.physicalSize = const Size(1100, 900);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);
    final gateway = MemoryAudioGateway();

    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          body: AudioConfigurationDialog(pc: 'PC-001', gateway: gateway),
        ),
      ),
    );
    await tester.pumpAndSettle();

    expect(find.text('Audio configuration · PC-001'), findsOneWidget);
    expect(find.text('Headphones'), findsWidgets);
    expect(find.byIcon(Icons.volume_up_rounded), findsNWidgets(4));

    await tester.tap(find.byKey(const Key('editAudioButton')));
    await tester.pump();
    await tester.tap(find.byKey(const Key('audioMute-PlaybackVolume')));
    final playbackDropdown = find.byKey(
      const Key('audioDevice-Playback:Default'),
    );
    await tester.ensureVisible(playbackDropdown);
    await tester.tap(playbackDropdown);
    await tester.pumpAndSettle();
    await tester.tap(find.text('HDMI').last);
    await tester.pumpAndSettle();
    await tester.tap(find.byKey(const Key('saveAudioButton')));
    await tester.pumpAndSettle();
    expect(gateway.saves, 1);
    expect(gateway.levelsJson, contains('"PlaybackMute": true'));
    expect(gateway.devicesJson, contains('"UnrecognizedField": "preserved"'));
    expect(
      AudioConfigurationData.fromJsonText(
        levelsJson: gateway.levelsJson,
        devicesJson: gateway.devicesJson,
      ).devices.firstWhere((device) => device['Name'] == 'HDMI')['Default'],
      isTrue,
    );

    await tester.tap(find.byKey(const Key('audioJsonButton')));
    await tester.pump();
    expect(find.byKey(const Key('audioLevelsJsonEditor')), findsOneWidget);
    expect(find.byKey(const Key('audioDevicesJsonEditor')), findsOneWidget);

    await tester.tap(find.byKey(const Key('audioLogButton')));
    await tester.pumpAndSettle();
    expect(find.text('AudioDeviceStartup.log · PC-001'), findsOneWidget);
    expect(find.textContaining('Audio restored'), findsOneWidget);
    final logText = tester.widget<SelectableText>(
      find.byKey(const Key('audioLogContents')),
    );
    final infoSpan = logText.textSpan!.children!
        .whereType<TextSpan>()
        .firstWhere((span) => span.text == 'INFO');
    expect(infoSpan.style?.color, const Color(0xff22c55e));
  });
}
