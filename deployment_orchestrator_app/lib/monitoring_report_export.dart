import 'dart:io';
import 'dart:typed_data';

import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;

class MonitoringExportEntry {
  const MonitoringExportEntry({
    required this.pc,
    required this.status,
    required this.lastScanned,
    required this.deploymentRecorded,
    required this.connectivity,
    required this.deployment,
    required this.audio,
    required this.display,
    required this.software,
  });

  final String pc;
  final String status;
  final String lastScanned;
  final String deploymentRecorded;
  final Map<String, String> connectivity;
  final Map<String, String> deployment;
  final Map<String, String> audio;
  final Map<String, String> display;
  final Map<String, String> software;

  Iterable<MapEntry<String, String>> get allFields sync* {
    yield* connectivity.entries;
    yield* deployment.entries;
    yield* audio.entries;
    yield* display.entries;
    yield* software.entries;
  }
}

class MonitoringExportReport {
  const MonitoringExportReport({
    required this.generatedAt,
    required this.scope,
    required this.entries,
  });

  final DateTime generatedAt;
  final String scope;
  final List<MonitoringExportEntry> entries;
}

String buildMonitoringCsv(MonitoringExportReport report) {
  final detailHeaders = <String>[];
  for (final entry in report.entries) {
    for (final field in entry.allFields) {
      if (!detailHeaders.contains(field.key)) detailHeaders.add(field.key);
    }
  }
  final headers = <String>[
    'Report generated',
    'Report scope',
    'PC',
    'Status',
    'Last scanned',
    'Deployment recorded',
    ...detailHeaders,
  ];
  final rows = <List<String>>[
    headers,
    for (final entry in report.entries)
      [
        _formatTimestamp(report.generatedAt),
        report.scope,
        entry.pc,
        entry.status,
        entry.lastScanned,
        entry.deploymentRecorded,
        for (final header in detailHeaders)
          entry.allFields
                  .where((field) => field.key == header)
                  .map((field) => field.value)
                  .firstOrNull ??
              '',
      ],
  ];
  return rows.map((row) => row.map(_csvCell).join(',')).join('\r\n');
}

String _csvCell(String value) {
  var safe = value;
  if (safe.isNotEmpty && '=+-@'.contains(safe[0])) safe = "'$safe";
  return '"${safe.replaceAll('"', '""')}"';
}

Future<Uint8List> buildMonitoringPdf(MonitoringExportReport report) async {
  final theme = await _loadPdfTheme();
  final document = pw.Document(
    title: 'Monitoring Scan Report',
    author: 'Windows Audio and Display Baseline Enforcer',
    subject: report.scope,
  );
  final healthyCount = report.entries
      .where((entry) => entry.status == 'Healthy')
      .length;
  final attentionCount = report.entries.length - healthyCount;

  document.addPage(
    pw.MultiPage(
      pageFormat: PdfPageFormat.letter,
      theme: theme,
      margin: const pw.EdgeInsets.fromLTRB(36, 44, 36, 42),
      header: (context) => context.pageNumber == 1
          ? pw.SizedBox()
          : pw.Container(
              padding: const pw.EdgeInsets.only(bottom: 8),
              decoration: const pw.BoxDecoration(
                border: pw.Border(
                  bottom: pw.BorderSide(color: PdfColors.grey400, width: 0.5),
                ),
              ),
              child: pw.Text(
                'Monitoring Scan Report',
                style: const pw.TextStyle(
                  color: PdfColors.grey700,
                  fontSize: 9,
                ),
              ),
            ),
      footer: (context) => pw.Align(
        alignment: pw.Alignment.centerRight,
        child: pw.Text(
          'Page ${context.pageNumber} of ${context.pagesCount}',
          style: const pw.TextStyle(color: PdfColors.grey600, fontSize: 8),
        ),
      ),
      build: (context) => [
        pw.Text(
          'Monitoring Scan Report',
          style: pw.TextStyle(
            color: PdfColors.blueGrey900,
            fontSize: 24,
            fontWeight: pw.FontWeight.bold,
          ),
        ),
        pw.SizedBox(height: 6),
        pw.Text(
          'Generated ${_formatTimestamp(report.generatedAt)}',
          style: const pw.TextStyle(color: PdfColors.grey700, fontSize: 10),
        ),
        pw.Text(
          report.scope,
          style: const pw.TextStyle(color: PdfColors.grey700, fontSize: 10),
        ),
        pw.SizedBox(height: 16),
        pw.Row(
          children: [
            _summaryCard('PCs', '${report.entries.length}', PdfColors.blue700),
            pw.SizedBox(width: 10),
            _summaryCard('Healthy', '$healthyCount', PdfColors.green700),
            pw.SizedBox(width: 10),
            _summaryCard('Attention', '$attentionCount', PdfColors.orange800),
          ],
        ),
        pw.SizedBox(height: 18),
        if (report.entries.isEmpty)
          pw.Text('No monitoring results were included in this report.'),
        for (final entry in report.entries) ...[
          _entryHeading(entry),
          _timestampTable(entry),
          ..._section('Connectivity', entry.connectivity),
          ..._section('Deployment and checks', entry.deployment),
          ..._section('Audio configuration', entry.audio),
          ..._section('Display configuration', entry.display),
          ..._section('Installed software', entry.software),
          pw.SizedBox(height: 14),
        ],
      ],
    ),
  );
  return document.save();
}

Future<pw.ThemeData?> _loadPdfTheme() async {
  const candidates = <(String, String)>[
    (r'C:\Windows\Fonts\segoeui.ttf', r'C:\Windows\Fonts\segoeuib.ttf'),
    (
      '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
      '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    ),
    (
      '/System/Library/Fonts/Supplemental/Arial.ttf',
      '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
    ),
  ];
  for (final (regularPath, boldPath) in candidates) {
    final regularFile = File(regularPath);
    final boldFile = File(boldPath);
    if (!await regularFile.exists() || !await boldFile.exists()) continue;
    final regularBytes = await regularFile.readAsBytes();
    final boldBytes = await boldFile.readAsBytes();
    return pw.ThemeData.withFont(
      base: pw.Font.ttf(regularBytes.buffer.asByteData()),
      bold: pw.Font.ttf(boldBytes.buffer.asByteData()),
    );
  }
  return null;
}

pw.Widget _summaryCard(String label, String value, PdfColor color) {
  return pw.Expanded(
    child: pw.Container(
      padding: const pw.EdgeInsets.all(10),
      decoration: pw.BoxDecoration(
        color: PdfColors.grey100,
        border: pw.Border(left: pw.BorderSide(color: color, width: 3)),
      ),
      child: pw.Column(
        crossAxisAlignment: pw.CrossAxisAlignment.start,
        children: [
          pw.Text(
            value,
            style: pw.TextStyle(
              color: color,
              fontSize: 16,
              fontWeight: pw.FontWeight.bold,
            ),
          ),
          pw.Text(label, style: const pw.TextStyle(fontSize: 9)),
        ],
      ),
    ),
  );
}

pw.Widget _entryHeading(MonitoringExportEntry entry) {
  final color = entry.status == 'Healthy'
      ? PdfColors.green700
      : entry.status == 'Offline' || entry.status == 'WinRM unavailable'
      ? PdfColors.red700
      : PdfColors.orange800;
  return pw.Container(
    margin: const pw.EdgeInsets.only(top: 6, bottom: 6),
    padding: const pw.EdgeInsets.symmetric(horizontal: 10, vertical: 8),
    decoration: pw.BoxDecoration(
      color: PdfColors.blueGrey50,
      border: pw.Border(left: pw.BorderSide(color: color, width: 4)),
    ),
    child: pw.Row(
      mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
      children: [
        pw.Text(
          entry.pc,
          style: pw.TextStyle(fontSize: 14, fontWeight: pw.FontWeight.bold),
        ),
        pw.Text(
          entry.status,
          style: pw.TextStyle(
            color: color,
            fontSize: 10,
            fontWeight: pw.FontWeight.bold,
          ),
        ),
      ],
    ),
  );
}

pw.Widget _timestampTable(MonitoringExportEntry entry) {
  return _fieldTable({
    'Last scanned': entry.lastScanned,
    'Deployment recorded': entry.deploymentRecorded,
  });
}

List<pw.Widget> _section(String title, Map<String, String> fields) {
  if (fields.isEmpty) return const [];
  return [
    pw.NewPage(freeSpace: 28 + (fields.length * 16)),
    pw.SizedBox(height: 7),
    pw.Text(
      title,
      style: pw.TextStyle(
        color: PdfColors.blueGrey800,
        fontSize: 10,
        fontWeight: pw.FontWeight.bold,
      ),
    ),
    pw.SizedBox(height: 3),
    _fieldTable(fields),
  ];
}

pw.Widget _fieldTable(Map<String, String> fields) {
  return pw.Table(
    border: pw.TableBorder.all(color: PdfColors.grey300, width: 0.4),
    columnWidths: const {
      0: pw.FlexColumnWidth(2.1),
      1: pw.FlexColumnWidth(4.9),
    },
    children: [
      for (final field in fields.entries)
        pw.TableRow(
          decoration: const pw.BoxDecoration(color: PdfColors.white),
          children: [
            pw.Padding(
              padding: const pw.EdgeInsets.all(4),
              child: pw.Text(
                field.key,
                style: pw.TextStyle(
                  fontSize: 8,
                  fontWeight: pw.FontWeight.bold,
                ),
              ),
            ),
            pw.Padding(
              padding: const pw.EdgeInsets.all(4),
              child: pw.Text(
                field.value.isEmpty ? 'Not available' : field.value,
                style: const pw.TextStyle(fontSize: 8),
              ),
            ),
          ],
        ),
    ],
  );
}

String _formatTimestamp(DateTime value) => value.toLocal().toIso8601String();
