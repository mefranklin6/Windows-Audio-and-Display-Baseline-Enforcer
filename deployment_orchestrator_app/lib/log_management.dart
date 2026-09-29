import 'dart:convert';
import 'dart:io';

import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

import 'log_formatting.dart';
import 'log_reporting.dart';

class LogStorageSummary {
  const LogStorageSummary({required this.bytes, required this.fileCount});

  final int bytes;
  final int fileCount;
}

class LogCleanupResult {
  const LogCleanupResult({
    required this.deletedFiles,
    required this.deletedBytes,
    required this.failedFiles,
  });

  final int deletedFiles;
  final int deletedBytes;
  final int failedFiles;
}

class LogFileInfo {
  const LogFileInfo({
    required this.file,
    required this.name,
    required this.size,
    required this.modified,
  });

  final File file;
  final String name;
  final int size;
  final DateTime modified;
}

class LogFileContents {
  const LogFileContents({required this.text, required this.truncated});

  final String text;
  final bool truncated;
}

typedef LogSummaryReader = Future<LogStorageSummary> Function(
  Directory directory,
);
typedef LogCleaner = Future<LogCleanupResult> Function(
  Directory directory, {
  DateTime? olderThan,
});

Future<List<File>> listLogFiles(Directory directory) async {
  if (!await directory.exists()) return const [];

  final files = <File>[];
  await for (final entity in directory.list(followLinks: false)) {
    if (entity is File && entity.path.toLowerCase().endsWith('.log')) {
      files.add(entity);
    }
  }
  return files;
}

Future<List<LogFileInfo>> readLogFileInfo(Directory directory) async {
  final files = await listLogFiles(directory);
  final logs = <LogFileInfo>[];
  for (final file in files) {
    try {
      final stat = await file.stat();
      logs.add(
        LogFileInfo(
          file: file,
          name: Uri.decodeComponent(file.uri.pathSegments.last),
          size: stat.size,
          modified: stat.modified,
        ),
      );
    } on FileSystemException {
      // A log may be removed between listing and inspection.
    }
  }
  logs.sort((left, right) => right.modified.compareTo(left.modified));
  return logs;
}

Future<LogFileContents> readLogFileContents(
  File file, {
  int maximumBytes = 2 * 1024 * 1024,
}) async {
  final length = await file.length();
  if (length <= maximumBytes) {
    return LogFileContents(text: await file.readAsString(), truncated: false);
  }

  final handle = await file.open();
  try {
    final start = length - maximumBytes;
    await handle.setPosition(start - 1);
    final startsAfterNewline = await handle.readByte() == 0x0a;
    await handle.setPosition(start);
    final bytes = await handle.read(maximumBytes);
    var text = utf8.decode(bytes, allowMalformed: true);
    if (!startsAfterNewline) {
      final firstNewline = text.indexOf('\n');
      if (firstNewline >= 0 && firstNewline < text.length - 1) {
        text = text.substring(firstNewline + 1);
      }
    }
    return LogFileContents(text: text, truncated: true);
  } finally {
    await handle.close();
  }
}

Future<int> exportLogFiles(
  Iterable<LogFileInfo> logs,
  Directory destination,
) async {
  final logList = logs.toList();
  if (logList.isNotEmpty &&
      _samePath(logList.first.file.parent.path, destination.path)) {
    throw const FileSystemException(
      'The export folder must be different from the log folder.',
    );
  }
  await destination.create(recursive: true);
  var exported = 0;
  for (final log in logList) {
    final destinationPath =
        '${destination.path}${Platform.pathSeparator}${log.name}';
    final exportFile = await _availableExportFile(File(destinationPath));
    await exportLogFile(log, exportFile);
    exported++;
  }
  return exported;
}

Future<File> _availableExportFile(File preferred) async {
  if (!await preferred.exists()) return preferred;
  final name = Uri.decodeComponent(preferred.uri.pathSegments.last);
  final dot = name.lastIndexOf('.');
  final base = dot > 0 ? name.substring(0, dot) : name;
  final extension = dot > 0 ? name.substring(dot) : '';
  for (var copy = 1; ; copy++) {
    final candidate = File(
      '${preferred.parent.path}${Platform.pathSeparator}$base ($copy)$extension',
    );
    if (!await candidate.exists()) return candidate;
  }
}

Future<void> exportLogFile(LogFileInfo log, File destination) async {
  if (_samePath(log.file.path, destination.path)) {
    throw const FileSystemException(
      'The export destination must be different from the original log.',
    );
  }
  await destination.parent.create(recursive: true);
  await log.file.copy(destination.path);
}

String formatLogTimestamp(DateTime value) {
  String twoDigits(int number) => number.toString().padLeft(2, '0');
  return '${value.year.toString().padLeft(4, '0')}-'
      '${twoDigits(value.month)}-${twoDigits(value.day)} '
      '${twoDigits(value.hour)}:${twoDigits(value.minute)}:'
      '${twoDigits(value.second)}';
}

bool _samePath(String left, String right) {
  final leftPath = File(left).absolute.path;
  final rightPath = File(right).absolute.path;
  return Platform.isWindows
      ? leftPath.toLowerCase() == rightPath.toLowerCase()
      : leftPath == rightPath;
}

Future<LogStorageSummary> readLogStorageSummary(Directory directory) async {
  final files = await listLogFiles(directory);
  var bytes = 0;
  for (final file in files) {
    try {
      bytes += await file.length();
    } on FileSystemException {
      // A log may be removed between listing and inspection.
    }
  }
  return LogStorageSummary(bytes: bytes, fileCount: files.length);
}

Future<LogCleanupResult> clearLogFiles(
  Directory directory, {
  DateTime? olderThan,
}) async {
  final files = await listLogFiles(directory);
  var deletedFiles = 0;
  var deletedBytes = 0;
  var failedFiles = 0;

  for (final file in files) {
    try {
      final stat = await file.stat();
      if (olderThan != null && !stat.modified.isBefore(olderThan)) continue;
      await file.delete();
      deletedFiles++;
      deletedBytes += stat.size;
    } on FileSystemException {
      failedFiles++;
    }
  }

  return LogCleanupResult(
    deletedFiles: deletedFiles,
    deletedBytes: deletedBytes,
    failedFiles: failedFiles,
  );
}

String formatFileSize(int bytes) {
  if (bytes < 1024) return '$bytes B';
  const units = ['KB', 'MB', 'GB', 'TB'];
  var value = bytes / 1024;
  var unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  final decimals = value >= 100 ? 0 : (value >= 10 ? 1 : 2);
  return '${value.toStringAsFixed(decimals)} ${units[unit]}';
}

enum LogRetentionUnit {
  minutes('minutes', 1),
  hours('hours', 60),
  days('days', 1440),
  weeks('weeks', 10080);

  const LogRetentionUnit(this.label, this.minutesPerUnit);

  final String label;
  final int minutesPerUnit;
}

enum LogRetentionSetting {
  forever('Forever'),
  days('Days'),
  months('Months'),
  years('Years');

  const LogRetentionSetting(this.label);

  final String label;
}

DateTime? logRetentionCutoff(
  DateTime now,
  int amount,
  LogRetentionSetting setting,
) {
  if (setting == LogRetentionSetting.forever || amount <= 0) return null;
  if (setting == LogRetentionSetting.days) {
    return now.subtract(Duration(days: amount));
  }
  if (setting == LogRetentionSetting.years) {
    return _clampedCalendarDate(now, now.year - amount, now.month);
  }
  final zeroBasedMonth = now.month - 1 - amount;
  final year = now.year + (zeroBasedMonth / 12).floor();
  final month = zeroBasedMonth - ((year - now.year) * 12) + 1;
  return _clampedCalendarDate(now, year, month);
}

DateTime _clampedCalendarDate(DateTime source, int year, int month) {
  final lastDay = DateTime(year, month + 1, 0).day;
  final day = source.day > lastDay ? lastDay : source.day;
  return DateTime(
    year,
    month,
    day,
    source.hour,
    source.minute,
    source.second,
    source.millisecond,
    source.microsecond,
  );
}

enum LogFileSort {
  newest('Newest first'),
  oldest('Oldest first'),
  name('File name'),
  largest('Largest first'),
  smallest('Smallest first');

  const LogFileSort(this.label);

  final String label;
}

class LogManagementSection extends StatefulWidget {
  const LogManagementSection({
    required this.logDirectory,
    this.summaryReader = readLogStorageSummary,
    this.enabled = true,
    this.showHeading = true,
    super.key,
  });

  final Directory logDirectory;
  final LogSummaryReader summaryReader;
  final bool enabled;
  final bool showHeading;

  @override
  State<LogManagementSection> createState() => _LogManagementSectionState();
}

class _LogManagementSectionState extends State<LogManagementSection> {
  LogStorageSummary? _summary;
  Object? _summaryError;
  bool _loading = false;

  @override
  void initState() {
    super.initState();
    _refresh();
  }

  @override
  void didUpdateWidget(LogManagementSection oldWidget) {
    super.didUpdateWidget(oldWidget);
    if (oldWidget.logDirectory.path != widget.logDirectory.path) _refresh();
  }

  Future<void> _refresh() async {
    if (_loading) return;
    setState(() {
      _loading = true;
      _summaryError = null;
    });
    try {
      final summary = await widget.summaryReader(widget.logDirectory);
      if (!mounted) return;
      setState(() => _summary = summary);
    } catch (error) {
      if (!mounted) return;
      setState(() => _summaryError = error);
    } finally {
      if (mounted) setState(() => _loading = false);
    }
  }

  Future<void> _showLogsDialog() async {
    try {
      final logs = await readLogFileInfo(widget.logDirectory);
      if (!mounted) return;
      List<LogReportEntry>? reportEntries;
      Object? reportError;
      var reportLoading = false;
      var mode = 0;
      var fileQuery = '';
      LogOperation? fileOperation;
      var fileSort = LogFileSort.newest;
      var reportQuery = '';
      var pcQuery = '';
      LogOperation? reportOperation;
      LogSeverity? reportSeverity;
      var reportPeriod = LogReportPeriod.all;
      var reportSort = LogReportSort.newest;
      await showDialog<void>(
        context: context,
        builder: (dialogContext) => StatefulBuilder(
          builder: (dialogContext, setDialogState) {
            final visibleLogs = logs.where((log) {
              if (fileOperation != null &&
                  logOperationForName(log.name) != fileOperation) {
                return false;
              }
              return fileQuery.trim().isEmpty ||
                  log.name.toLowerCase().contains(
                    fileQuery.trim().toLowerCase(),
                  );
            }).toList();
            visibleLogs.sort(switch (fileSort) {
              LogFileSort.newest => (left, right) => right.modified.compareTo(
                left.modified,
              ),
              LogFileSort.oldest => (left, right) => left.modified.compareTo(
                right.modified,
              ),
              LogFileSort.name => (
                left,
                right,
              ) => left.name.toLowerCase().compareTo(right.name.toLowerCase()),
              LogFileSort.largest => (left, right) => right.size.compareTo(
                left.size,
              ),
              LogFileSort.smallest => (left, right) => left.size.compareTo(
                right.size,
              ),
            });
            final allReportEntries = reportEntries ?? const <LogReportEntry>[];
            final filteredEntries = filterLogReportEntries(
              allReportEntries,
              query: reportQuery,
              pcQuery: pcQuery,
              operation: reportOperation,
              severity: reportSeverity,
              period: reportPeriod,
              sort: reportSort,
            );
            final reportScope = logReportScope(
              query: reportQuery,
              pcQuery: pcQuery,
              operation: reportOperation,
              severity: reportSeverity,
              period: reportPeriod,
            );
            final logsByName = <String, LogFileInfo>{
              for (final log in logs) log.name: log,
            };
            return AlertDialog(
              title: const Row(
                children: [
                  Icon(Icons.article_outlined),
                  SizedBox(width: 10),
                  Text('Log explorer'),
                ],
              ),
              content: SizedBox(
                width: 980,
                height: (MediaQuery.sizeOf(dialogContext).height - 180).clamp(
                  420.0,
                  760.0,
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    SegmentedButton<int>(
                      key: const Key('logExplorerMode'),
                      segments: const [
                        ButtonSegment(
                          value: 0,
                          icon: Icon(Icons.folder_outlined),
                          label: Text('Log files'),
                        ),
                        ButtonSegment(
                          value: 1,
                          icon: Icon(Icons.assessment_outlined),
                          label: Text('Report builder'),
                        ),
                      ],
                      selected: {mode},
                      onSelectionChanged: (selection) async {
                        final selectedMode = selection.first;
                        setDialogState(() => mode = selectedMode);
                        if (selectedMode != 1 ||
                            reportEntries != null ||
                            reportLoading) {
                          return;
                        }
                        setDialogState(() {
                          reportLoading = true;
                          reportError = null;
                        });
                        try {
                          final loaded = await readLogReportEntries(
                            logs.map((log) => log.file),
                          );
                          if (!dialogContext.mounted) return;
                          setDialogState(() => reportEntries = loaded);
                        } on Object catch (error) {
                          if (!dialogContext.mounted) return;
                          setDialogState(() => reportError = error);
                        } finally {
                          if (dialogContext.mounted) {
                            setDialogState(() => reportLoading = false);
                          }
                        }
                      },
                    ),
                    const SizedBox(height: 12),
                    if (mode == 0)
                      Expanded(
                        child: _buildLogFilesBrowser(
                          logs: visibleLogs,
                          totalCount: logs.length,
                          query: fileQuery,
                          operation: fileOperation,
                          sort: fileSort,
                          onQueryChanged: (value) =>
                              setDialogState(() => fileQuery = value),
                          onOperationChanged: (value) =>
                              setDialogState(() => fileOperation = value),
                          onSortChanged: (value) =>
                              setDialogState(() => fileSort = value),
                        ),
                      )
                    else
                      Expanded(
                        child: reportLoading
                            ? const Center(
                                child: Column(
                                  mainAxisSize: MainAxisSize.min,
                                  children: [
                                    Icon(Icons.hourglass_top, size: 32),
                                    SizedBox(height: 10),
                                    Text('Reading log entries…'),
                                  ],
                                ),
                              )
                            : reportError != null
                            ? Center(
                                child: Text(
                                  'Could not build the report: $reportError',
                                  textAlign: TextAlign.center,
                                ),
                              )
                            : reportEntries == null
                            ? const SizedBox.shrink()
                            : _buildLogReportBuilder(
                                entries: filteredEntries,
                                totalCount: allReportEntries.length,
                                logsByName: logsByName,
                                query: reportQuery,
                                pcQuery: pcQuery,
                                operation: reportOperation,
                                severity: reportSeverity,
                                period: reportPeriod,
                                sort: reportSort,
                                onQueryChanged: (value) =>
                                    setDialogState(() => reportQuery = value),
                                onPcChanged: (value) =>
                                    setDialogState(() => pcQuery = value),
                                onOperationChanged: (value) => setDialogState(
                                  () => reportOperation = value,
                                ),
                                onSeverityChanged: (value) => setDialogState(
                                  () => reportSeverity = value,
                                ),
                                onPeriodChanged: (value) =>
                                    setDialogState(() => reportPeriod = value),
                                onSortChanged: (value) =>
                                    setDialogState(() => reportSort = value),
                              ),
                      ),
                  ],
                ),
              ),
              actions: [
                if (mode == 1)
                  OutlinedButton.icon(
                    key: const Key('exportCustomLogReportButton'),
                    onPressed: reportEntries == null || filteredEntries.isEmpty
                        ? null
                        : () => _exportCustomLogReport(
                            filteredEntries,
                            reportScope,
                          ),
                    icon: const Icon(Icons.download_outlined),
                    label: Text(
                      'Export ${filteredEntries.length} ${filteredEntries.length == 1 ? 'entry' : 'entries'}',
                    ),
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
    } on Object catch (error) {
      if (!mounted) return;
      _showMessage('Could not open the log list: $error', isError: true);
    }
  }

  Widget _buildLogFilesBrowser({
    required List<LogFileInfo> logs,
    required int totalCount,
    required String query,
    required LogOperation? operation,
    required LogFileSort sort,
    required ValueChanged<String> onQueryChanged,
    required ValueChanged<LogOperation?> onOperationChanged,
    required ValueChanged<LogFileSort> onSortChanged,
  }) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        Wrap(
          spacing: 10,
          runSpacing: 10,
          children: [
            SizedBox(
              width: 330,
              child: TextFormField(
                key: const Key('logFileSearchField'),
                initialValue: query,
                onChanged: onQueryChanged,
                decoration: const InputDecoration(
                  labelText: 'Search file names',
                  prefixIcon: Icon(Icons.search),
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
              ),
            ),
            SizedBox(
              width: 200,
              child: DropdownButtonFormField<LogOperation?>(
                key: const Key('logFileOperationFilter'),
                initialValue: operation,
                decoration: const InputDecoration(
                  labelText: 'Operation',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  const DropdownMenuItem<LogOperation?>(
                    value: null,
                    child: Text('All operations'),
                  ),
                  for (final value in LogOperation.values)
                    DropdownMenuItem<LogOperation?>(
                      value: value,
                      child: Text(value.label),
                    ),
                ],
                onChanged: onOperationChanged,
              ),
            ),
            SizedBox(
              width: 190,
              child: DropdownButtonFormField<LogFileSort>(
                key: const Key('logFileSort'),
                initialValue: sort,
                decoration: const InputDecoration(
                  labelText: 'Sort',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  for (final value in LogFileSort.values)
                    DropdownMenuItem(value: value, child: Text(value.label)),
                ],
                onChanged: (value) {
                  if (value != null) onSortChanged(value);
                },
              ),
            ),
          ],
        ),
        const SizedBox(height: 10),
        Text(
          'Showing ${logs.length} of $totalCount log files',
          style: Theme.of(context).textTheme.bodySmall,
        ),
        const SizedBox(height: 6),
        Expanded(
          child: logs.isEmpty
              ? const Center(child: Text('No log files match these filters.'))
              : ListView.separated(
                  itemCount: logs.length,
                  separatorBuilder: (_, _) => const Divider(height: 1),
                  itemBuilder: (context, index) {
                    final log = logs[index];
                    final operation = logOperationForName(log.name);
                    return ListTile(
                      key: Key('logFile-${log.name}'),
                      leading: Icon(_logOperationIcon(operation)),
                      title: Text(
                        log.name,
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                      ),
                      subtitle: Text(
                        '${operation.label} · '
                        '${formatLogTimestamp(log.modified)} · '
                        '${formatFileSize(log.size)}',
                      ),
                      trailing: Wrap(
                        spacing: 4,
                        children: [
                          IconButton(
                            key: Key('exportLog-${log.name}'),
                            tooltip: 'Download ${log.name}',
                            onPressed: () => _exportLog(log),
                            icon: const Icon(Icons.download_outlined),
                          ),
                          TextButton.icon(
                            key: Key('viewLog-${log.name}'),
                            onPressed: () => _showLogViewer(log),
                            icon: const Icon(Icons.visibility_outlined),
                            label: const Text('View in app'),
                          ),
                        ],
                      ),
                      onTap: () => _showLogViewer(log),
                    );
                  },
                ),
        ),
      ],
    );
  }

  Widget _buildLogReportBuilder({
    required List<LogReportEntry> entries,
    required int totalCount,
    required Map<String, LogFileInfo> logsByName,
    required String query,
    required String pcQuery,
    required LogOperation? operation,
    required LogSeverity? severity,
    required LogReportPeriod period,
    required LogReportSort sort,
    required ValueChanged<String> onQueryChanged,
    required ValueChanged<String> onPcChanged,
    required ValueChanged<LogOperation?> onOperationChanged,
    required ValueChanged<LogSeverity?> onSeverityChanged,
    required ValueChanged<LogReportPeriod> onPeriodChanged,
    required ValueChanged<LogReportSort> onSortChanged,
  }) {
    final displayedEntries = entries.take(1000).toList();
    final pcCount = entries
        .map((entry) => entry.pc)
        .whereType<String>()
        .toSet()
        .length;
    final issueCount = entries
        .where((entry) => entry.severity.rank >= LogSeverity.warning.rank)
        .length;
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        Wrap(
          spacing: 10,
          runSpacing: 10,
          children: [
            SizedBox(
              width: 330,
              child: TextFormField(
                key: const Key('logReportSearchField'),
                initialValue: query,
                onChanged: onQueryChanged,
                decoration: const InputDecoration(
                  labelText: 'Search messages and source logs',
                  prefixIcon: Icon(Icons.search),
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
              ),
            ),
            SizedBox(
              width: 190,
              child: TextFormField(
                key: const Key('logReportPcField'),
                initialValue: pcQuery,
                onChanged: onPcChanged,
                decoration: const InputDecoration(
                  labelText: 'PC contains',
                  prefixIcon: Icon(Icons.computer_outlined),
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
              ),
            ),
            SizedBox(
              width: 190,
              child: DropdownButtonFormField<LogOperation?>(
                key: const Key('logReportOperationFilter'),
                initialValue: operation,
                decoration: const InputDecoration(
                  labelText: 'Operation',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  const DropdownMenuItem<LogOperation?>(
                    value: null,
                    child: Text('All operations'),
                  ),
                  for (final value in LogOperation.values)
                    DropdownMenuItem<LogOperation?>(
                      value: value,
                      child: Text(value.label),
                    ),
                ],
                onChanged: onOperationChanged,
              ),
            ),
            SizedBox(
              width: 170,
              child: DropdownButtonFormField<LogSeverity?>(
                key: const Key('logReportSeverityFilter'),
                initialValue: severity,
                decoration: const InputDecoration(
                  labelText: 'Severity',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  const DropdownMenuItem<LogSeverity?>(
                    value: null,
                    child: Text('All severities'),
                  ),
                  for (final value in LogSeverity.values)
                    DropdownMenuItem<LogSeverity?>(
                      value: value,
                      child: Text(value.label),
                    ),
                ],
                onChanged: onSeverityChanged,
              ),
            ),
            SizedBox(
              width: 170,
              child: DropdownButtonFormField<LogReportPeriod>(
                key: const Key('logReportPeriodFilter'),
                initialValue: period,
                decoration: const InputDecoration(
                  labelText: 'Time window',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  for (final value in LogReportPeriod.values)
                    DropdownMenuItem(value: value, child: Text(value.label)),
                ],
                onChanged: (value) {
                  if (value != null) onPeriodChanged(value);
                },
              ),
            ),
            SizedBox(
              width: 170,
              child: DropdownButtonFormField<LogReportSort>(
                key: const Key('logReportSort'),
                initialValue: sort,
                decoration: const InputDecoration(
                  labelText: 'Sort',
                  border: OutlineInputBorder(),
                  isDense: true,
                ),
                items: [
                  for (final value in LogReportSort.values)
                    DropdownMenuItem(value: value, child: Text(value.label)),
                ],
                onChanged: (value) {
                  if (value != null) onSortChanged(value);
                },
              ),
            ),
          ],
        ),
        const SizedBox(height: 10),
        Wrap(
          spacing: 8,
          runSpacing: 6,
          children: [
            Chip(label: Text('${entries.length} of $totalCount entries')),
            Chip(label: Text('$pcCount PCs')),
            Chip(label: Text('$issueCount warnings or errors')),
            if (entries.length > displayedEntries.length)
              Chip(
                label: Text(
                  'Previewing first ${displayedEntries.length}; export includes all',
                ),
              ),
          ],
        ),
        const SizedBox(height: 6),
        if (displayedEntries.isNotEmpty)
          Container(
            padding: const EdgeInsets.symmetric(vertical: 7, horizontal: 4),
            color: Theme.of(context).colorScheme.surfaceContainerHighest,
            child: const Row(
              children: [
                SizedBox(width: 145, child: Text('Timestamp')),
                SizedBox(width: 92, child: Text('Operation')),
                SizedBox(width: 76, child: Text('Severity')),
                SizedBox(width: 120, child: Text('PC')),
                Expanded(child: Text('Message')),
                SizedBox(width: 88, child: Text('Source')),
              ],
            ),
          ),
        Expanded(
          child: displayedEntries.isEmpty
              ? const Center(child: Text('No entries match this report.'))
              : ListView.separated(
                  itemCount: displayedEntries.length,
                  separatorBuilder: (_, _) => const Divider(height: 1),
                  itemBuilder: (context, index) {
                    final entry = displayedEntries[index];
                    final sourceLog = logsByName[entry.sourceName];
                    return Padding(
                      padding: const EdgeInsets.symmetric(
                        vertical: 7,
                        horizontal: 4,
                      ),
                      child: Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          SizedBox(
                            width: 145,
                            child: Text(
                              formatLogTimestamp(entry.timestamp),
                              style: Theme.of(context).textTheme.bodySmall,
                            ),
                          ),
                          SizedBox(
                            width: 92,
                            child: Text(
                              entry.operation.label,
                              style: Theme.of(context).textTheme.bodySmall,
                            ),
                          ),
                          SizedBox(
                            width: 76,
                            child: Text(
                              entry.severity.label,
                              style: TextStyle(
                                color: _logSeverityColor(entry.severity),
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                          ),
                          SizedBox(
                            width: 120,
                            child: Text(
                              entry.pc ?? '—',
                              overflow: TextOverflow.ellipsis,
                            ),
                          ),
                          Expanded(
                            child: Tooltip(
                              message:
                                  '${entry.sourceName}:${entry.lineNumber}',
                              child: Text(entry.message),
                            ),
                          ),
                          SizedBox(
                            width: 88,
                            child: sourceLog == null
                                ? const SizedBox.shrink()
                                : Row(
                                    mainAxisAlignment: MainAxisAlignment.end,
                                    children: [
                                      IconButton(
                                        key: Key(
                                          'exportReportSource-${entry.sourceName}-${entry.lineNumber}',
                                        ),
                                        tooltip: 'Export ${entry.sourceName}',
                                        constraints:
                                            const BoxConstraints.tightFor(
                                              width: 40,
                                              height: 40,
                                            ),
                                        padding: EdgeInsets.zero,
                                        onPressed: () => _exportLog(sourceLog),
                                        icon: const Icon(
                                          Icons.download_outlined,
                                          size: 19,
                                        ),
                                      ),
                                      IconButton(
                                        key: Key(
                                          'viewReportSource-${entry.sourceName}-${entry.lineNumber}',
                                        ),
                                        tooltip: 'View ${entry.sourceName}',
                                        constraints:
                                            const BoxConstraints.tightFor(
                                              width: 40,
                                              height: 40,
                                            ),
                                        padding: EdgeInsets.zero,
                                        onPressed: () =>
                                            _showLogViewer(sourceLog),
                                        icon: const Icon(
                                          Icons.open_in_new,
                                          size: 19,
                                        ),
                                      ),
                                    ],
                                  ),
                          ),
                        ],
                      ),
                    );
                  },
                ),
        ),
      ],
    );
  }

  IconData _logOperationIcon(LogOperation operation) {
    return switch (operation) {
      LogOperation.deployment => Icons.rocket_launch_outlined,
      LogOperation.monitoring => Icons.monitor_heart_outlined,
      LogOperation.uninstall => Icons.delete_sweep_outlined,
    };
  }

  Color _logSeverityColor(LogSeverity severity) {
    final dark = Theme.of(context).brightness == Brightness.dark;
    return switch (severity) {
      LogSeverity.debug =>
        dark ? const Color(0xffc4c7c5) : const Color(0xff5f6368),
      LogSeverity.info =>
        dark ? const Color(0xff6dd58c) : const Color(0xff146c2e),
      LogSeverity.warning =>
        dark ? const Color(0xffffb95c) : const Color(0xff8a4a00),
      LogSeverity.error || LogSeverity.fatal =>
        dark ? const Color(0xffffb4ab) : const Color(0xffb3261e),
    };
  }

  Future<void> _showLogViewer(LogFileInfo log) async {
    try {
      final contents = await readLogFileContents(log.file);
      if (!mounted) return;
      final entries = parseLogEntries(log.name, contents.text);
      var rawMode = entries.isEmpty;
      var query = '';
      var pcQuery = '';
      LogSeverity? severity;
      var sort = LogReportSort.oldest;
      await showDialog<void>(
        context: context,
        builder: (dialogContext) => StatefulBuilder(
          builder: (dialogContext, setDialogState) {
            final filteredEntries = filterLogReportEntries(
              entries,
              query: query,
              pcQuery: pcQuery,
              severity: severity,
              sort: sort,
            );
            return AlertDialog(
              title: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(log.name, overflow: TextOverflow.ellipsis),
                  const SizedBox(height: 4),
                  Text(
                    '${formatLogTimestamp(log.modified)} · '
                    '${formatFileSize(log.size)}',
                    style: Theme.of(dialogContext).textTheme.bodySmall,
                  ),
                ],
              ),
              content: SizedBox(
                width: 980,
                height: (MediaQuery.sizeOf(dialogContext).height - 180).clamp(
                  380.0,
                  760.0,
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    if (contents.truncated)
                      Container(
                        key: const Key('logPreviewTruncatedNotice'),
                        margin: const EdgeInsets.only(bottom: 8),
                        padding: const EdgeInsets.all(10),
                        decoration: BoxDecoration(
                          color: Theme.of(dialogContext)
                              .colorScheme
                              .secondaryContainer,
                          borderRadius: BorderRadius.circular(8),
                        ),
                        child: const Text(
                          'This is a large log. Showing only the most recent 2 MB.',
                        ),
                      ),
                    if (entries.isNotEmpty) ...[
                      SegmentedButton<bool>(
                        key: const Key('logViewerMode'),
                        segments: const [
                          ButtonSegment(
                            value: false,
                            icon: Icon(Icons.table_rows_outlined),
                            label: Text('Filtered'),
                          ),
                          ButtonSegment(
                            value: true,
                            icon: Icon(Icons.code),
                            label: Text('Raw'),
                          ),
                        ],
                        selected: {rawMode},
                        onSelectionChanged: (selection) =>
                            setDialogState(() => rawMode = selection.first),
                      ),
                      const SizedBox(height: 10),
                    ],
                    if (!rawMode && entries.isNotEmpty) ...[
                      Wrap(
                        spacing: 10,
                        runSpacing: 10,
                        children: [
                          SizedBox(
                            width: 310,
                            child: TextFormField(
                              key: const Key('logViewerSearchField'),
                              initialValue: query,
                              onChanged: (value) =>
                                  setDialogState(() => query = value),
                              decoration: const InputDecoration(
                                labelText: 'Search this log',
                                prefixIcon: Icon(Icons.search),
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                            ),
                          ),
                          SizedBox(
                            width: 180,
                            child: TextFormField(
                              key: const Key('logViewerPcField'),
                              initialValue: pcQuery,
                              onChanged: (value) =>
                                  setDialogState(() => pcQuery = value),
                              decoration: const InputDecoration(
                                labelText: 'PC contains',
                                prefixIcon: Icon(Icons.computer_outlined),
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                            ),
                          ),
                          SizedBox(
                            width: 170,
                            child: DropdownButtonFormField<LogSeverity?>(
                              key: const Key('logViewerSeverityFilter'),
                              initialValue: severity,
                              decoration: const InputDecoration(
                                labelText: 'Severity',
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                              items: [
                                const DropdownMenuItem<LogSeverity?>(
                                  value: null,
                                  child: Text('All severities'),
                                ),
                                for (final value in LogSeverity.values)
                                  DropdownMenuItem<LogSeverity?>(
                                    value: value,
                                    child: Text(value.label),
                                  ),
                              ],
                              onChanged: (value) =>
                                  setDialogState(() => severity = value),
                            ),
                          ),
                          SizedBox(
                            width: 170,
                            child: DropdownButtonFormField<LogReportSort>(
                              key: const Key('logViewerSort'),
                              initialValue: sort,
                              decoration: const InputDecoration(
                                labelText: 'Sort',
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                              items: [
                                for (final value in LogReportSort.values)
                                  DropdownMenuItem(
                                    value: value,
                                    child: Text(value.label),
                                  ),
                              ],
                              onChanged: (value) {
                                if (value != null) {
                                  setDialogState(() => sort = value);
                                }
                              },
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 8),
                      Text(
                        'Showing ${filteredEntries.length} of ${entries.length} entries',
                        key: const Key('logViewerResultCount'),
                        style: Theme.of(dialogContext).textTheme.bodySmall,
                      ),
                      const SizedBox(height: 6),
                    ],
                    Expanded(
                      child: rawMode || entries.isEmpty
                          ? Container(
                              padding: const EdgeInsets.all(12),
                              decoration: BoxDecoration(
                                color: Theme.of(dialogContext)
                                    .colorScheme
                                    .surfaceContainerHighest,
                                borderRadius: BorderRadius.circular(8),
                              ),
                              child: SingleChildScrollView(
                                child: SelectableText.rich(
                                  TextSpan(
                                    children: buildLogSeveritySpans(
                                      contents.text,
                                    ),
                                  ),
                                  key: const Key('logViewerContents'),
                                  style: const TextStyle(
                                    fontFamily: 'monospace',
                                    fontSize: 13,
                                    height: 1.35,
                                  ),
                                ),
                              ),
                            )
                          : filteredEntries.isEmpty
                          ? const Center(
                              child: Text(
                                'No entries match the current filters.',
                              ),
                            )
                          : ListView.separated(
                              key: const Key('filteredLogViewerEntries'),
                              itemCount: filteredEntries.length,
                              separatorBuilder: (_, _) =>
                                  const Divider(height: 1),
                              itemBuilder: (context, index) {
                                final entry = filteredEntries[index];
                                final color = _logSeverityColor(entry.severity);
                                return Padding(
                                  padding: const EdgeInsets.symmetric(
                                    vertical: 8,
                                    horizontal: 4,
                                  ),
                                  child: Row(
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      SizedBox(
                                        width: 145,
                                        child: Text(
                                          formatLogTimestamp(entry.timestamp),
                                          style: Theme.of(context)
                                              .textTheme
                                              .bodySmall,
                                        ),
                                      ),
                                      SizedBox(
                                        width: 76,
                                        child: Text(
                                          entry.severity.label,
                                          style: TextStyle(
                                            color: color,
                                            fontWeight: FontWeight.w700,
                                          ),
                                        ),
                                      ),
                                      SizedBox(
                                        width: 125,
                                        child: Text(
                                          entry.pc ?? '—',
                                          overflow: TextOverflow.ellipsis,
                                        ),
                                      ),
                                      Expanded(
                                        child: SelectableText(entry.message),
                                      ),
                                    ],
                                  ),
                                );
                              },
                            ),
                    ),
                  ],
                ),
              ),
              actions: [
                OutlinedButton.icon(
                  key: const Key('exportOpenLogButton'),
                  onPressed: () => _exportLog(log),
                  icon: const Icon(Icons.download_outlined),
                  label: const Text('Download'),
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
    } on Object catch (error) {
      if (!mounted) return;
      _showMessage('Could not open ${log.name}: $error', isError: true);
    }
  }

  Future<void> _exportLog(LogFileInfo log) async {
    try {
      final location = await getSaveLocation(
        suggestedName: log.name,
        acceptedTypeGroups: const [
          XTypeGroup(label: 'Log file', extensions: ['log']),
        ],
      );
      if (location == null) return;
      var path = location.path;
      if (!path.toLowerCase().endsWith('.log')) path = '$path.log';
      await exportLogFile(log, File(path));
      if (!mounted) return;
      _showMessage('Exported ${log.name} to $path.');
    } on Object catch (error) {
      if (!mounted) return;
      _showMessage('Could not export ${log.name}: $error', isError: true);
    }
  }

  Future<void> _exportCustomLogReport(
    List<LogReportEntry> entries,
    String scope,
  ) async {
    try {
      final now = DateTime.now();
      final stamp =
          '${now.year.toString().padLeft(4, '0')}'
          '${now.month.toString().padLeft(2, '0')}'
          '${now.day.toString().padLeft(2, '0')}-'
          '${now.hour.toString().padLeft(2, '0')}'
          '${now.minute.toString().padLeft(2, '0')}';
      final location = await getSaveLocation(
        suggestedName: 'log-report-$stamp.csv',
        acceptedTypeGroups: const [
          XTypeGroup(label: 'CSV report', extensions: ['csv']),
        ],
      );
      if (location == null) return;
      var path = location.path;
      if (!path.toLowerCase().endsWith('.csv')) path = '$path.csv';
      await File(path).writeAsString(
        '\uFEFF${buildLogReportCsv(entries, generatedAt: now, scope: scope)}',
      );
      if (!mounted) return;
      _showMessage('Exported ${entries.length} report entries to $path.');
    } on Object catch (error) {
      if (!mounted) return;
      _showMessage('Could not export the log report: $error', isError: true);
    }
  }

  Future<void> _exportAllLogs() async {
    try {
      final logs = await readLogFileInfo(widget.logDirectory);
      if (!mounted) return;
      if (logs.isEmpty) {
        _showMessage('There are no logs to export.');
        return;
      }
      final destination = await getDirectoryPath(
        initialDirectory: widget.logDirectory.parent.path,
        confirmButtonText: 'Export logs here',
      );
      if (destination == null) return;
      final exported = await exportLogFiles(logs, Directory(destination));
      if (!mounted) return;
      _showMessage(
        'Exported $exported log ${exported == 1 ? 'file' : 'files'} to '
        '$destination.',
      );
    } on Object catch (error) {
      if (!mounted) return;
      _showMessage('Could not export logs: $error', isError: true);
    }
  }

  void _showMessage(String message, {bool isError = false}) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: isError ? Theme.of(context).colorScheme.error : null,
      ),
    );
  }

  Future<void> _showClearDialog() async {
    var clearEverything = false;
    var amount = '30';
    var unit = LogRetentionUnit.days;
    String? validationError;
    var clearing = false;

    await showDialog<void>(
      context: context,
      barrierDismissible: false,
      builder: (dialogContext) => StatefulBuilder(
        builder: (_, setDialogState) {
          Future<void> clear() async {
            final value = double.tryParse(amount.trim());
            if (!clearEverything && (value == null || value <= 0)) {
              setDialogState(
                () => validationError = 'Enter a number greater than zero.',
              );
              return;
            }

            setDialogState(() {
              clearing = true;
              validationError = null;
            });
            final cutoff = clearEverything
                ? null
                : DateTime.now().subtract(
                    Duration(
                      seconds: (value! * unit.minutesPerUnit * 60).round(),
                    ),
                  );
            final result = await clearLogFiles(
              widget.logDirectory,
              olderThan: cutoff,
            );
            if (!dialogContext.mounted) return;
            Navigator.of(dialogContext).pop();
            await _refresh();
            if (!mounted) return;
            final failure = result.failedFiles == 0
                ? ''
                : ' ${result.failedFiles} could not be removed.';
            ScaffoldMessenger.of(context).showSnackBar(
              SnackBar(
                content: Text(
                  'Removed ${result.deletedFiles} log ${result.deletedFiles == 1 ? 'file' : 'files'} '
                  '(${formatFileSize(result.deletedBytes)}).$failure',
                ),
              ),
            );
          }

          return AlertDialog(
            title: const Text('Clear logs'),
            content: SizedBox(
              width: 430,
              child: RadioGroup<bool>(
                groupValue: clearEverything,
                onChanged: (value) {
                  if (!clearing) {
                    setDialogState(() => clearEverything = value ?? false);
                  }
                },
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    RadioListTile<bool>(
                      key: const Key('clearOlderLogsOption'),
                      contentPadding: EdgeInsets.zero,
                      title: const Text('Clear logs older than'),
                      value: false,
                      enabled: !clearing,
                    ),
                    Row(
                      children: [
                        Expanded(
                          child: TextFormField(
                            key: const Key('clearLogAgeAmountField'),
                            initialValue: amount,
                            enabled: !clearEverything && !clearing,
                            keyboardType: const TextInputType.numberWithOptions(
                              decimal: true,
                            ),
                            inputFormatters: [
                              FilteringTextInputFormatter.allow(
                                RegExp(r'^\d*\.?\d*$'),
                              ),
                            ],
                            onChanged: (value) => amount = value,
                            decoration: InputDecoration(
                              labelText: 'Age',
                              errorText: validationError,
                              border: const OutlineInputBorder(),
                            ),
                          ),
                        ),
                        const SizedBox(width: 12),
                        Expanded(
                          child: DropdownButtonFormField<LogRetentionUnit>(
                            key: const Key('clearLogAgeUnitField'),
                            initialValue: unit,
                            decoration: const InputDecoration(
                              labelText: 'Unit',
                              border: OutlineInputBorder(),
                            ),
                            items: LogRetentionUnit.values
                                .map(
                                  (value) => DropdownMenuItem(
                                    value: value,
                                    child: Text(value.label),
                                  ),
                                )
                                .toList(),
                            onChanged: clearEverything || clearing
                                ? null
                                : (value) => setDialogState(
                                    () => unit = value ?? unit,
                                  ),
                          ),
                        ),
                      ],
                    ),
                    RadioListTile<bool>(
                      key: const Key('clearAllLogsOption'),
                      contentPadding: EdgeInsets.zero,
                      title: const Text('Clear all logs'),
                      subtitle: const Text('This cannot be undone.'),
                      value: true,
                      enabled: !clearing,
                    ),
                    const Text(
                      'Deployment records and configuration backups are not removed.',
                    ),
                  ],
                ),
              ),
            ),
            actions: [
              TextButton(
                onPressed: clearing
                    ? null
                    : () => Navigator.of(dialogContext).pop(),
                child: const Text('Cancel'),
              ),
              FilledButton.icon(
                key: const Key('confirmClearLogsButton'),
                onPressed: clearing ? null : clear,
                icon: clearing
                    ? const Icon(Icons.hourglass_top, size: 18)
                    : const Icon(Icons.delete_outline),
                label: Text(clearing ? 'Clearing…' : 'Clear logs'),
              ),
            ],
          );
        },
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final summary = _summary;
    final details = _summaryError != null
        ? 'Could not read the log directory.'
        : summary == null
        ? 'Calculating log size…'
        : '${formatFileSize(summary.bytes)} across ${summary.fileCount} '
              'log ${summary.fileCount == 1 ? 'file' : 'files'}';

    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        if (widget.showHeading) ...[
          Text('Log storage', style: Theme.of(context).textTheme.titleMedium),
          const SizedBox(height: 8),
        ],
        ListTile(
          key: const Key('logStorageSummary'),
          contentPadding: EdgeInsets.zero,
          leading: const Icon(Icons.description_outlined),
          title: Text(details),
          subtitle: Text(widget.logDirectory.path),
          trailing: IconButton(
            key: const Key('refreshLogSizeButton'),
            tooltip: 'Refresh log size',
            onPressed: _loading ? null : _refresh,
            icon: _loading
                ? const Icon(Icons.hourglass_top)
                : const Icon(Icons.refresh),
          ),
        ),
        Wrap(
          alignment: WrapAlignment.end,
          spacing: 8,
          runSpacing: 8,
          children: [
            OutlinedButton.icon(
              key: const Key('viewLogsButton'),
              onPressed: widget.enabled && !_loading ? _showLogsDialog : null,
              icon: const Icon(Icons.article_outlined),
              label: const Text('View logs'),
            ),
            OutlinedButton.icon(
              key: const Key('exportAllLogsButton'),
              onPressed: widget.enabled && !_loading ? _exportAllLogs : null,
              icon: const Icon(Icons.download_outlined),
              label: const Text('Export all…'),
            ),
            OutlinedButton.icon(
              key: const Key('clearLogsButton'),
              onPressed: widget.enabled && !_loading ? _showClearDialog : null,
              icon: const Icon(Icons.delete_outline),
              label: const Text('Clear logs…'),
            ),
          ],
        ),
      ],
    );
  }
}
