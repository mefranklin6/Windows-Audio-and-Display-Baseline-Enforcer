import 'dart:convert';
import 'dart:io';

const applicationVersion = String.fromEnvironment(
  'APP_VERSION',
  defaultValue: '1.0.0',
);
const releasesPageUrl =
    'https://github.com/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases';
const latestReleaseApiUrl =
    'https://api.github.com/repos/mefranklin6/Windows-Audio-and-Display-Baseline-Enforcer/releases/latest';

class UpdateCheckResult {
  const UpdateCheckResult({
    required this.currentVersion,
    required this.latestVersion,
    required this.releaseUrl,
    required this.downloadUrl,
  });

  final String currentVersion;
  final String latestVersion;
  final String releaseUrl;
  final String? downloadUrl;

  bool get updateAvailable =>
      compareVersions(latestVersion, currentVersion) > 0;

  String get preferredUrl => downloadUrl ?? releaseUrl;
}

Future<UpdateCheckResult> checkForUpdates() async {
  final client = HttpClient();
  try {
    final request = await client.getUrl(Uri.parse(latestReleaseApiUrl));
    request.headers
      ..set(HttpHeaders.acceptHeader, 'application/vnd.github+json')
      ..set(HttpHeaders.userAgentHeader, 'Windows-Audio-Display-Enforcer');
    final response = await request.close();
    final body = await response.transform(utf8.decoder).join();
    if (response.statusCode != HttpStatus.ok) {
      throw HttpException(
        'GitHub returned HTTP ${response.statusCode}.',
        uri: Uri.parse(latestReleaseApiUrl),
      );
    }

    final decoded = jsonDecode(body);
    if (decoded is! Map<String, dynamic>) {
      throw const FormatException('GitHub returned an invalid release record.');
    }
    final latestVersion = (decoded['tag_name'] as String? ?? '').trim();
    if (latestVersion.isEmpty) {
      throw const FormatException('The latest release has no version tag.');
    }
    final releaseUrl =
        (decoded['html_url'] as String?)?.trim().isNotEmpty == true
        ? (decoded['html_url'] as String).trim()
        : releasesPageUrl;
    String? downloadUrl;
    final assets = decoded['assets'];
    if (assets is List) {
      for (final asset in assets.whereType<Map<String, dynamic>>()) {
        final name = (asset['name'] as String? ?? '').toLowerCase();
        final url = (asset['browser_download_url'] as String? ?? '').trim();
        if (name.endsWith('-setup.exe') && url.isNotEmpty) {
          downloadUrl = url;
          break;
        }
      }
    }
    return UpdateCheckResult(
      currentVersion: applicationVersion,
      latestVersion: latestVersion,
      releaseUrl: releaseUrl,
      downloadUrl: downloadUrl,
    );
  } finally {
    client.close(force: true);
  }
}

Future<void> openWebUrl(String url) async {
  if (!Platform.isWindows) {
    throw UnsupportedError('Opening release links is supported on Windows.');
  }
  final result = await Process.run('explorer.exe', [url]);
  if (result.exitCode != 0) {
    throw ProcessException(
      'explorer.exe',
      [url],
      '${result.stderr}',
      result.exitCode,
    );
  }
}

int compareVersions(String left, String right) {
  final leftParts = _versionParts(left);
  final rightParts = _versionParts(right);
  final count = leftParts.length > rightParts.length
      ? leftParts.length
      : rightParts.length;
  for (var index = 0; index < count; index++) {
    final leftPart = index < leftParts.length ? leftParts[index] : 0;
    final rightPart = index < rightParts.length ? rightParts[index] : 0;
    if (leftPart != rightPart) return leftPart.compareTo(rightPart);
  }
  return 0;
}

List<int> _versionParts(String version) {
  final normalized = version.trim().replaceFirst(RegExp(r'^[^0-9]*'), '');
  final core = normalized.split(RegExp(r'[-+]')).first;
  return core
      .split('.')
      .map((part) => int.tryParse(part) ?? 0)
      .toList(growable: false);
}
