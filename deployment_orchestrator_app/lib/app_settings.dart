import 'dart:convert';
import 'dart:io';

abstract interface class SettingsStore {
  Future<Map<String, dynamic>?> load();

  Future<void> save(Map<String, dynamic> settings);
}

class JsonSettingsStore implements SettingsStore {
  JsonSettingsStore(this.file);

  factory JsonSettingsStore.forCurrentUser() {
    final applicationData =
        Platform.environment['APPDATA'] ??
        Platform.environment['LOCALAPPDATA'] ??
        Directory.current.absolute.path;
    return JsonSettingsStore(
      File(
        '$applicationData${Platform.pathSeparator}'
        'Windows Audio and Display Baseline Enforcer Orchestrator${Platform.pathSeparator}'
        'settings.json',
      ),
    );
  }

  final File file;

  @override
  Future<Map<String, dynamic>?> load() async {
    try {
      final decoded = jsonDecode(await file.readAsString());
      return decoded is Map<String, dynamic> ? decoded : null;
    } on FileSystemException {
      return null;
    } on FormatException {
      return null;
    }
  }

  @override
  Future<void> save(Map<String, dynamic> settings) async {
    await file.parent.create(recursive: true);
    await file.writeAsString(
      const JsonEncoder.withIndent('  ').convert(settings),
      flush: true,
    );
  }
}
