import 'dart:convert';
import 'dart:io';

enum LogOperation {
  deployment('Deployment'),
  monitoring('Monitoring'),
  uninstall('Uninstall');

  const LogOperation(this.label);

  final String label;
}

enum LogSeverity {
  debug('Debug', 0),
  info('Info', 1),
  warning('Warning', 2),
  error('Error', 3),
  fatal('Fatal', 4);

  const LogSeverity(this.label, this.rank);

  final String label;
  final int rank;
}

enum LogReportSort {
  newest('Newest first'),
  oldest('Oldest first'),
  pc('PC name'),
  severity('Severity');

  const LogReportSort(this.label);

  final String label;
}

enum LogReportPeriod {
  all('All time', null),
  last24Hours('Last 24 hours', Duration(hours: 24)),
  last7Days('Last 7 days', Duration(days: 7)),
  last30Days('Last 30 days', Duration(days: 30));

  const LogReportPeriod(this.label, this.duration);

  final String label;
  final Duration? duration;
}

class LogReportEntry {
  const LogReportEntry({
    required this.timestamp,
    required this.severity,
    required this.operation,
    required this.message,
    required this.sourceName,
    required this.lineNumber,
    this.pc,
  });

  final DateTime timestamp;
  final LogSeverity severity;
  final LogOperation operation;
  final String? pc;
  final String message;
  final String sourceName;
  final int lineNumber;
}

LogOperation logOperationForName(String name) {
  final normalized = name.toLowerCase();
  if (normalized.startsWith('monitor-')) return LogOperation.monitoring;
  if (normalized.startsWith('uninstall-')) return LogOperation.uninstall;
  return LogOperation.deployment;
}

List<LogReportEntry> parseLogEntries(String sourceName, String contents) {
  final operation = logOperationForName(sourceName);
  final result = <LogReportEntry>[];
  final pattern = RegExp(
    r'^(\S+)\s+(DEBUG|INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\s+(.*)$',
    caseSensitive: false,
  );
  final pcPattern = RegExp(r'^([^:\s]+):\s*(.*)$');
  final lines = const LineSplitter().convert(contents);

  for (var index = 0; index < lines.length; index++) {
    final match = pattern.firstMatch(lines[index]);
    if (match == null) continue;
    final timestamp = DateTime.tryParse(match.group(1)!);
    if (timestamp == null) continue;
    final body = match.group(3)!;
    final pcMatch = pcPattern.firstMatch(body);
    result.add(
      LogReportEntry(
        timestamp: timestamp,
        severity: _parseSeverity(match.group(2)!),
        operation: operation,
        pc: pcMatch?.group(1),
        message: pcMatch?.group(2) ?? body,
        sourceName: sourceName,
        lineNumber: index + 1,
      ),
    );
  }
  return result;
}

Future<List<LogReportEntry>> readLogReportEntries(Iterable<File> files) async {
  final entries = <LogReportEntry>[];
  for (final file in files) {
    try {
      final name = Uri.decodeComponent(file.uri.pathSegments.last);
      final contents = utf8.decode(
        await file.readAsBytes(),
        allowMalformed: true,
      );
      entries.addAll(parseLogEntries(name, contents));
    } on FileSystemException {
      // A log may be removed while the report is being assembled.
    }
  }
  return entries;
}

List<LogReportEntry> filterLogReportEntries(
  Iterable<LogReportEntry> entries, {
  String query = '',
  String pcQuery = '',
  LogOperation? operation,
  LogSeverity? severity,
  LogReportPeriod period = LogReportPeriod.all,
  LogReportSort sort = LogReportSort.newest,
  DateTime? now,
}) {
  final normalizedQuery = query.trim().toLowerCase();
  final normalizedPc = pcQuery.trim().toLowerCase();
  final duration = period.duration;
  final cutoff = duration == null
      ? null
      : (now ?? DateTime.now()).subtract(duration);
  final filtered = entries.where((entry) {
    if (operation != null && entry.operation != operation) return false;
    if (severity != null && entry.severity != severity) return false;
    if (cutoff != null && entry.timestamp.isBefore(cutoff)) return false;
    if (normalizedPc.isNotEmpty &&
        !(entry.pc ?? '').toLowerCase().contains(normalizedPc)) {
      return false;
    }
    if (normalizedQuery.isNotEmpty) {
      final searchable = <String>[
        entry.message,
        entry.pc ?? '',
        entry.sourceName,
        entry.severity.label,
        entry.operation.label,
      ].join('\n').toLowerCase();
      if (!searchable.contains(normalizedQuery)) return false;
    }
    return true;
  }).toList();

  filtered.sort(switch (sort) {
    LogReportSort.newest => (left, right) => right.timestamp.compareTo(
      left.timestamp,
    ),
    LogReportSort.oldest => (left, right) => left.timestamp.compareTo(
      right.timestamp,
    ),
    LogReportSort.pc => (left, right) {
      final comparison = (left.pc ?? '').toLowerCase().compareTo(
        (right.pc ?? '').toLowerCase(),
      );
      return comparison != 0
          ? comparison
          : right.timestamp.compareTo(left.timestamp);
    },
    LogReportSort.severity => (left, right) {
      final comparison = right.severity.rank.compareTo(left.severity.rank);
      return comparison != 0
          ? comparison
          : right.timestamp.compareTo(left.timestamp);
    },
  });
  return filtered;
}

String buildLogReportCsv(
  Iterable<LogReportEntry> entries, {
  required DateTime generatedAt,
  required String scope,
}) {
  final rows = <List<String>>[
    const [
      'Report generated',
      'Report scope',
      'Timestamp',
      'Operation',
      'Severity',
      'PC',
      'Message',
      'Source log',
      'Source line',
    ],
    for (final entry in entries)
      [
        generatedAt.toIso8601String(),
        scope,
        entry.timestamp.toIso8601String(),
        entry.operation.label,
        entry.severity.label,
        entry.pc ?? '',
        entry.message,
        entry.sourceName,
        entry.lineNumber.toString(),
      ],
  ];
  return rows.map((row) => row.map(_csvCell).join(',')).join('\r\n');
}

String logReportScope({
  required String query,
  required String pcQuery,
  required LogOperation? operation,
  required LogSeverity? severity,
  required LogReportPeriod period,
}) {
  return <String>[
    period.label,
    operation?.label ?? 'All operations',
    severity?.label ?? 'All severities',
    if (pcQuery.trim().isNotEmpty) 'PC contains “${pcQuery.trim()}”',
    if (query.trim().isNotEmpty) 'Text contains “${query.trim()}”',
  ].join('; ');
}

LogSeverity _parseSeverity(String value) {
  return switch (value.toUpperCase()) {
    'DEBUG' => LogSeverity.debug,
    'WARNING' || 'WARN' => LogSeverity.warning,
    'ERROR' => LogSeverity.error,
    'FATAL' || 'CRITICAL' => LogSeverity.fatal,
    _ => LogSeverity.info,
  };
}

String _csvCell(String value) {
  var safe = value;
  if (safe.isNotEmpty && '=+-@'.contains(safe[0])) safe = "'$safe";
  return '"${safe.replaceAll('"', '""')}"';
}
