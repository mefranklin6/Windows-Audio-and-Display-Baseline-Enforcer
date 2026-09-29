import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/log_reporting.dart';

void main() {
  const contents = '''2026-09-29T09:48:16.058672 INFO PC-001: Monitoring started
2026-09-29T09:48:18.098044 WARNING PC-002: Display is missing
2026-09-29T09:48:19.100000 ERROR Global failure
not a structured log line''';

  test('parses structured log fields and operation type', () {
    final entries = parseLogEntries('monitor-2026.log', contents);

    expect(entries, hasLength(3));
    expect(entries.first.operation, LogOperation.monitoring);
    expect(entries.first.pc, 'PC-001');
    expect(entries.first.message, 'Monitoring started');
    expect(entries[1].severity, LogSeverity.warning);
    expect(entries.last.pc, isNull);
    expect(entries.last.lineNumber, 3);
  });

  test('filters and sorts report entries', () {
    final entries = parseLogEntries('monitor-2026.log', contents);

    final filtered = filterLogReportEntries(
      entries,
      query: 'display',
      pcQuery: '002',
      severity: LogSeverity.warning,
      period: LogReportPeriod.last7Days,
      now: DateTime.parse('2026-09-30T00:00:00'),
    );

    expect(filtered, hasLength(1));
    expect(filtered.single.pc, 'PC-002');
  });

  test('builds spreadsheet-safe CSV for the filtered report', () {
    final entry = LogReportEntry(
      timestamp: DateTime.parse('2026-09-29T09:48:19'),
      severity: LogSeverity.error,
      operation: LogOperation.deployment,
      pc: 'PC-001',
      message: '=unsafe formula',
      sourceName: 'run.log',
      lineNumber: 4,
    );

    final csv = buildLogReportCsv(
      [entry],
      generatedAt: DateTime.parse('2026-09-29T10:00:00'),
      scope: 'Errors only',
    );

    expect(csv, contains('"\'=unsafe formula"'));
    expect(csv, contains('"Errors only"'));
    expect(csv, contains('"run.log"'));
  });
}
