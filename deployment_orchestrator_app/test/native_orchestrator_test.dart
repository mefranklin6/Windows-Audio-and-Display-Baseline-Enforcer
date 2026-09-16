import 'dart:convert';
import 'dart:io';

import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/native_orchestrator.dart';

class FakeCommandExecutor implements CommandExecutor {
  FakeCommandExecutor(this.handler);

  final Future<CommandResult> Function(String, List<String>) handler;
  final List<List<String>> calls = [];
  bool _cancelled = false;

  @override
  bool get isCancelled => _cancelled;

  @override
  void cancel() => _cancelled = true;

  @override
  Future<CommandResult> run(
    String executable,
    List<String> arguments, {
    required String workingDirectory,
  }) async {
    calls.add([executable, ...arguments]);
    return handler(executable, arguments);
  }
}

void main() {
  late Directory projectRoot;

  setUp(() async {
    projectRoot = await Directory.systemTemp.createTemp('native-orchestrator-');
    await Directory(
      '${projectRoot.path}${Platform.pathSeparator}installer_scripts',
    ).create();
    await Directory(
      '${projectRoot.path}${Platform.pathSeparator}utility_scripts',
    ).create();
  });

  tearDown(() async {
    if (projectRoot.existsSync()) {
      await projectRoot.delete(recursive: true);
    }
  });

  test(
    'deploys enabled scripts in order and records structured warnings',
    () async {
      final logs = <String>[];
      final executor = FakeCommandExecutor((_, arguments) async {
        if (arguments.first == 'ping' || arguments.first == 'Invoke-Command') {
          return const CommandResult(exitCode: 0);
        }
        if (arguments.any(
          (value) => value.endsWith('InstallAudioDeviceCmdlets.ps1'),
        )) {
          return const CommandResult(
            exitCode: 0,
            stdout: 'INFO: Audio installed\nWARNING: Restart recommended\n',
          );
        }
        return const CommandResult(
          exitCode: 0,
          stdout: 'INFO: Cleanup complete',
        );
      });
      final orchestrator = NativeOrchestrator(
        projectRoot: projectRoot.path,
        maxWorkers: 2,
        executor: executor,
        onLog: logs.add,
        clock: () => DateTime.utc(2026, 9, 16, 12),
      );

      final report = await orchestrator.deploy(
        ['PC-001', 'PC-001'],
        const DeploymentOptions(
          audioRecall: true,
          displayRecall: false,
          bgInfoInstall: false,
          desktopShortcuts: false,
          bgInfoFolder: '',
        ),
      );

      final pcs = report['pcs'] as List<dynamic>;
      expect(pcs, hasLength(1));
      final pc = pcs.single as Map<String, dynamic>;
      expect(pc['highest_severity'], 'warning');
      final scripts = pc['scripts'] as List<dynamic>;
      expect(scripts.map((item) => (item as Map<String, dynamic>)['name']), [
        r'installer_scripts\InstallAudioDeviceCmdlets.ps1',
        r'installer_scripts\Cleanup.ps1',
      ]);
      expect((scripts.first as Map<String, dynamic>)['messages'], [
        'Restart recommended',
      ]);
      expect(
        logs.any((line) => line.contains('PC-001: Deployment complete')),
        isTrue,
      );

      final record = File(
        '${projectRoot.path}${Platform.pathSeparator}logs'
        '${Platform.pathSeparator}deployment_records'
        '${Platform.pathSeparator}pc-001.json',
      );
      final intent =
          jsonDecode(await record.readAsString()) as Map<String, dynamic>;
      expect(intent['audio_recall'], isTrue);
      expect(intent['display_recall'], isFalse);
    },
  );

  test('reports a failed ping without running installer scripts', () async {
    final executor = FakeCommandExecutor(
      (_, _) async => const CommandResult(exitCode: 1),
    );
    final orchestrator = NativeOrchestrator(
      projectRoot: projectRoot.path,
      maxWorkers: 1,
      executor: executor,
    );

    final report = await orchestrator.deploy(
      ['OFFLINE-PC'],
      const DeploymentOptions(
        audioRecall: true,
        displayRecall: true,
        bgInfoInstall: false,
        desktopShortcuts: true,
        bgInfoFolder: '',
      ),
    );

    final pc = (report['pcs'] as List<dynamic>).single as Map<String, dynamic>;
    expect(pc['highest_severity'], 'warning');
    expect(pc['scripts'], isEmpty);
    expect(
      (pc['issues'] as List<dynamic>).single,
      containsPair('message', 'Ping test failed'),
    );
    expect(executor.calls, hasLength(1));
  });

  test(
    'monitor parses PowerShell JSON and includes deployment intent',
    () async {
      final progress = <String>[];
      final executor = FakeCommandExecutor((_, _) async {
        return const CommandResult(
          exitCode: 0,
          stdout:
              'diagnostic output\n'
              '{"pc":"PC-002","online":true,"winrm":true,'
              '"cts_deployed":true,"audio_configured":true,'
              '"display_configured":true,"logout_shortcut":true,'
              '"reboot_shortcut":true,"bginfo_deployed":false,'
              '"bginfo_executable":false,"bginfo_profile":false,'
              '"bginfo_background":false,"bginfo_startup":false,'
              '"bginfo_startup_method":"",'
              '"audio_device_cmdlets_versions":["3.3"],'
              '"display_config_versions":["6.0.1"]}',
        );
      });
      final deployment = NativeOrchestrator(
        projectRoot: projectRoot.path,
        maxWorkers: 1,
        executor: FakeCommandExecutor((_, arguments) async {
          if (arguments.first == 'ping' ||
              arguments.first == 'Invoke-Command') {
            return const CommandResult(exitCode: 0);
          }
          return const CommandResult(exitCode: 0);
        }),
      );
      await deployment.deploy(
        ['PC-002'],
        const DeploymentOptions(
          audioRecall: true,
          displayRecall: true,
          bgInfoInstall: false,
          desktopShortcuts: true,
          bgInfoFolder: '',
        ),
      );

      final monitor = NativeOrchestrator(
        projectRoot: projectRoot.path,
        maxWorkers: 1,
        executor: executor,
        onMonitoringProgress: progress.add,
      );
      final report = await monitor.monitor(['PC-002']);

      final pc =
          (report['pcs'] as List<dynamic>).single as Map<String, dynamic>;
      expect(pc['online'], isTrue);
      expect(pc['audio_device_cmdlets_versions'], ['3.3']);
      expect(
        (pc['deployment_intent'] as Map<String, dynamic>)['desktop_shortcuts'],
        isTrue,
      );
      expect(progress, ['PC-002']);
    },
  );

  test('monitor returns a complete failure shape for invalid output', () async {
    final executor = FakeCommandExecutor(
      (_, _) async => const CommandResult(exitCode: 0, stdout: 'not json'),
    );
    final orchestrator = NativeOrchestrator(
      projectRoot: projectRoot.path,
      maxWorkers: 1,
      executor: executor,
    );

    final report = await orchestrator.monitor(['PC-003']);
    final pc = (report['pcs'] as List<dynamic>).single as Map<String, dynamic>;

    expect(pc['online'], isFalse);
    expect(pc['winrm'], isFalse);
    expect(pc['error'], startsWith('Invalid PowerShell monitoring result:'));
    expect(pc['audio_device_cmdlets_versions'], isEmpty);
  });

  test('honors the worker limit while preserving target order', () async {
    var active = 0;
    var maximumActive = 0;
    final executor = FakeCommandExecutor((_, arguments) async {
      active++;
      if (active > maximumActive) maximumActive = active;
      await Future<void>.delayed(const Duration(milliseconds: 10));
      active--;
      final pc = arguments.last;
      return CommandResult(
        exitCode: 0,
        stdout: '{"pc":"$pc","online":true,"winrm":true}',
      );
    });
    final orchestrator = NativeOrchestrator(
      projectRoot: projectRoot.path,
      maxWorkers: 2,
      executor: executor,
    );

    final report = await orchestrator.monitor([
      'PC-1',
      'PC-2',
      'PC-3',
      'PC-4',
      'PC-5',
    ]);
    final pcs = (report['pcs'] as List<dynamic>)
        .map((result) => (result as Map<String, dynamic>)['pc'])
        .toList();

    expect(maximumActive, 2);
    expect(pcs, ['PC-1', 'PC-2', 'PC-3', 'PC-4', 'PC-5']);
  });
}
