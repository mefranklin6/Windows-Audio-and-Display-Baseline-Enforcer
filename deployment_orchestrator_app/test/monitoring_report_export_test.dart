import 'dart:convert';

import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/monitoring_report_export.dart';

void main() {
  final report = MonitoringExportReport(
    generatedAt: DateTime.utc(2026, 9, 18, 16, 30),
    scope: 'All PCs',
    entries: const [
      MonitoringExportEntry(
        pc: 'PC-001',
        status: 'Healthy',
        lastScanned: '2026-09-18T16:29:00.000Z',
        deploymentRecorded: '2026-09-17T20:00:00.000Z',
        connectivity: {
          'Online': 'yes',
          'WinRM available': 'yes',
          'Monitoring error': '',
        },
        deployment: {
          'CTS deployment': 'Present',
          'BGInfo deployment': 'Present',
        },
        audio: {
          'Playback volume': '55%',
          'Playback mute': 'Unknown',
          'Default playback device': 'Headphones, Room A',
        },
        display: {
          'Display mode': 'Extended desktop',
          'Display details': '#1 Display: 1920x1080',
        },
        software: {
          'AudioDeviceCmdlets versions': '3.3',
          'DisplayConfig versions': '6.0.1',
        },
      ),
    ],
  );

  test('builds a flat, escaped CSV with scan and deployment timestamps', () {
    final csv = buildMonitoringCsv(report);

    expect(csv, startsWith('"Report generated","Report scope","PC"'));
    expect(csv, contains('"Last scanned"'));
    expect(csv, contains('"Deployment recorded"'));
    expect(csv, contains('"Playback mute"'));
    expect(csv, contains('"Headphones, Room A"'));
    expect(csv, contains('"2026-09-18T16:29:00.000Z"'));
    expect(csv.split('\r\n'), hasLength(2));
  });

  test('guards spreadsheet formulas in exported cells', () {
    final unsafe = MonitoringExportReport(
      generatedAt: report.generatedAt,
      scope: report.scope,
      entries: [
        MonitoringExportEntry(
          pc: '=HYPERLINK("bad")',
          status: 'Attention needed',
          lastScanned: '',
          deploymentRecorded: '',
          connectivity: const {},
          deployment: const {},
          audio: const {},
          display: const {},
          software: const {},
        ),
      ],
    );

    expect(buildMonitoringCsv(unsafe), contains('"\'=HYPERLINK(""bad"")"'));
  });

  test('builds a valid PDF containing report pages', () async {
    final bytes = await buildMonitoringPdf(report);

    expect(ascii.decode(bytes.take(5).toList()), '%PDF-');
    expect(bytes.length, greaterThan(1000));
  });
}
