import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:deployment_orchestrator_app/log_management.dart';

void main() {
  late Directory temporaryDirectory;

  setUp(() async {
    temporaryDirectory = await Directory.systemTemp.createTemp(
      'log-management-',
    );
  });

  tearDown(() async {
    if (await temporaryDirectory.exists()) {
      await temporaryDirectory.delete(recursive: true);
    }
  });

  test('summarizes only top-level log files', () async {
    await File('${temporaryDirectory.path}${Platform.pathSeparator}run.log')
        .writeAsString('12345');
    await File('${temporaryDirectory.path}${Platform.pathSeparator}record.json')
        .writeAsString('not a log');
    final records = await Directory(
      '${temporaryDirectory.path}${Platform.pathSeparator}deployment_records',
    ).create();
    await File('${records.path}${Platform.pathSeparator}nested.log')
        .writeAsString('preserved');

    final summary = await readLogStorageSummary(temporaryDirectory);

    expect(summary.fileCount, 1);
    expect(summary.bytes, 5);
  });

  test('clears only logs older than the selected cutoff', () async {
    final oldLog = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}old.log',
    ).writeAsString('old');
    final recentLog = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}recent.log',
    ).writeAsString('recent');
    final record = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}record.json',
    ).writeAsString('record');
    await oldLog.setLastModified(
      DateTime.now().subtract(const Duration(days: 31)),
    );

    final result = await clearLogFiles(
      temporaryDirectory,
      olderThan: DateTime.now().subtract(const Duration(days: 30)),
    );

    expect(result.deletedFiles, 1);
    expect(result.deletedBytes, 3);
    expect(await oldLog.exists(), isFalse);
    expect(await recentLog.exists(), isTrue);
    expect(await record.exists(), isTrue);
  });

  test('clears every log without removing non-log files', () async {
    final firstLog = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}first.log',
    ).writeAsString('first');
    final secondLog = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}second.LOG',
    ).writeAsString('second');
    final record = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}record.json',
    ).writeAsString('keep');

    final result = await clearLogFiles(temporaryDirectory);

    expect(result.deletedFiles, 2);
    expect(await firstLog.exists(), isFalse);
    expect(await secondLog.exists(), isFalse);
    expect(await record.exists(), isTrue);
  });

  test('lists logs newest first with file metadata', () async {
    final older = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}older.log',
    ).writeAsString('old');
    final newer = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}newer.log',
    ).writeAsString('newer');
    await older.setLastModified(DateTime(2026, 1, 1));
    await newer.setLastModified(DateTime(2026, 2, 1));

    final logs = await readLogFileInfo(temporaryDirectory);

    expect(logs.map((log) => log.name), ['newer.log', 'older.log']);
    expect(logs.first.size, 5);
  });

  test('reads only the newest complete lines of a large log', () async {
    final log = await File(
      '${temporaryDirectory.path}${Platform.pathSeparator}large.log',
    ).writeAsString('discard this line\nkeep one\nkeep two\n');

    final contents = await readLogFileContents(log, maximumBytes: 18);

    expect(contents.truncated, isTrue);
    expect(contents.text, 'keep one\nkeep two\n');
  });

  test('exports logs without changing their contents', () async {
    await File('${temporaryDirectory.path}${Platform.pathSeparator}source.log')
        .writeAsString('export me');
    final info = (await readLogFileInfo(temporaryDirectory)).single;
    final destination = Directory(
      '${temporaryDirectory.path}${Platform.pathSeparator}exported',
    );

    final count = await exportLogFiles([info], destination);

    expect(count, 1);
    expect(
      await File('${destination.path}${Platform.pathSeparator}source.log')
          .readAsString(),
      'export me',
    );
  });

  test('bulk export does not overwrite an existing log', () async {
    await File('${temporaryDirectory.path}${Platform.pathSeparator}source.log')
        .writeAsString('new contents');
    final info = (await readLogFileInfo(temporaryDirectory)).single;
    final destination = await Directory(
      '${temporaryDirectory.path}${Platform.pathSeparator}exported',
    ).create();
    final existing = await File(
      '${destination.path}${Platform.pathSeparator}source.log',
    ).writeAsString('existing contents');

    await exportLogFiles([info], destination);

    expect(await existing.readAsString(), 'existing contents');
    expect(
      await File('${destination.path}${Platform.pathSeparator}source (1).log')
          .readAsString(),
      'new contents',
    );
  });

  testWidgets('shows size and offers complete or age-based cleanup', (
    tester,
  ) async {
    await tester.pumpWidget(
      MaterialApp(
        home: Scaffold(
          body: LogManagementSection(
            logDirectory: temporaryDirectory,
            summaryReader: (_) async =>
                const LogStorageSummary(bytes: 1025, fileCount: 2),
          ),
        ),
      ),
    );
    await tester.pumpAndSettle();

    expect(find.text('1.00 KB across 2 log files'), findsOneWidget);
    expect(find.byKey(const Key('viewLogsButton')), findsOneWidget);
    expect(find.byKey(const Key('exportAllLogsButton')), findsOneWidget);
    await tester.tap(find.byKey(const Key('clearLogsButton')));
    await tester.pump();

    expect(find.byKey(const Key('clearOlderLogsOption')), findsOneWidget);
    expect(find.byKey(const Key('clearLogAgeAmountField')), findsOneWidget);
    expect(find.byKey(const Key('clearLogAgeUnitField')), findsOneWidget);
    expect(find.byKey(const Key('clearAllLogsOption')), findsOneWidget);
    expect(find.byKey(const Key('confirmClearLogsButton')), findsOneWidget);
  });

  test('formats byte sizes for display', () {
    expect(formatFileSize(0), '0 B');
    expect(formatFileSize(1024), '1.00 KB');
    expect(formatFileSize(10 * 1024), '10.0 KB');
    expect(formatFileSize(2 * 1024 * 1024), '2.00 MB');
  });

  test('calculates calendar-aware retention cutoffs', () {
    final now = DateTime(2026, 3, 31, 14, 30);

    expect(
      logRetentionCutoff(now, 1, LogRetentionSetting.months),
      DateTime(2026, 2, 28, 14, 30),
    );
    expect(
      logRetentionCutoff(
        DateTime(2024, 2, 29, 8),
        1,
        LogRetentionSetting.years,
      ),
      DateTime(2023, 2, 28, 8),
    );
    expect(
      logRetentionCutoff(now, 30, LogRetentionSetting.days),
      now.subtract(const Duration(days: 30)),
    );
    expect(logRetentionCutoff(now, 30, LogRetentionSetting.forever), isNull);
  });
}
