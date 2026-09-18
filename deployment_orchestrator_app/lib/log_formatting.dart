import 'package:flutter/material.dart';

List<TextSpan> buildLogSeveritySpans(String text) {
  final severityPattern = RegExp(
    r'\b(INFO|WARNING|WARN|ERROR|FATAL|CRITICAL)\b',
    caseSensitive: false,
  );
  final spans = <TextSpan>[];
  var cursor = 0;

  for (final match in severityPattern.allMatches(text)) {
    if (match.start > cursor) {
      spans.add(TextSpan(text: text.substring(cursor, match.start)));
    }
    final severity = match.group(0)!;
    final normalized = severity.toUpperCase();
    final color = switch (normalized) {
      'INFO' => const Color(0xff22c55e),
      'WARNING' || 'WARN' => const Color(0xffff9800),
      'ERROR' || 'FATAL' || 'CRITICAL' => const Color(0xffef4444),
      _ => null,
    };
    spans.add(
      TextSpan(
        text: severity,
        style: TextStyle(
          color: color,
          fontWeight: normalized == 'FATAL' || normalized == 'CRITICAL'
              ? FontWeight.bold
              : null,
        ),
      ),
    );
    cursor = match.end;
  }

  if (cursor < text.length) {
    spans.add(TextSpan(text: text.substring(cursor)));
  }
  return spans;
}
