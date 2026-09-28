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
    this.checksumUrl,
    this.installerFileName,
  });

  final String currentVersion;
  final String latestVersion;
  final String releaseUrl;
  final String? downloadUrl;
  final String? checksumUrl;
  final String? installerFileName;

  bool get updateAvailable =>
      compareVersions(latestVersion, currentVersion) > 0;

  String get preferredUrl => downloadUrl ?? releaseUrl;

  bool get canInstall =>
      updateAvailable &&
      downloadUrl != null &&
      checksumUrl != null &&
      installerFileName != null;
}

class PreparedUpdate {
  const PreparedUpdate({required this.installer, required this.sha256});

  final File installer;
  final String sha256;
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

    return parseRelease(jsonDecode(body));
  } finally {
    client.close(force: true);
  }
}

UpdateCheckResult parseRelease(Object? decoded) {
  if (decoded is! Map<String, dynamic>) {
    throw const FormatException('GitHub returned an invalid release record.');
  }
  final latestVersion = (decoded['tag_name'] as String? ?? '').trim();
  if (latestVersion.isEmpty) {
    throw const FormatException('The latest release has no version tag.');
  }
  final releaseUrl = (decoded['html_url'] as String?)?.trim().isNotEmpty == true
      ? (decoded['html_url'] as String).trim()
      : releasesPageUrl;
  String? downloadUrl;
  String? checksumUrl;
  String? installerFileName;
  final assetUrls = <String, String>{};
  final assets = decoded['assets'];
  if (assets is List) {
    for (final asset in assets.whereType<Map<String, dynamic>>()) {
      final name = (asset['name'] as String? ?? '').trim();
      final url = (asset['browser_download_url'] as String? ?? '').trim();
      if (name.isNotEmpty && url.isNotEmpty) {
        assetUrls[name.toLowerCase()] = url;
      }
      if (name.toLowerCase().endsWith('-setup.exe') && url.isNotEmpty) {
        installerFileName ??= name;
        downloadUrl ??= url;
      }
    }
  }
  if (installerFileName != null) {
    if (installerFileName.contains('/') ||
        installerFileName.contains(r'\') ||
        installerFileName == '.' ||
        installerFileName == '..') {
      throw const FormatException('The release installer has an invalid name.');
    }
    checksumUrl = assetUrls['${installerFileName.toLowerCase()}.sha256'];
  }
  return UpdateCheckResult(
    currentVersion: applicationVersion,
    latestVersion: latestVersion,
    releaseUrl: releaseUrl,
    downloadUrl: downloadUrl,
    checksumUrl: checksumUrl,
    installerFileName: installerFileName,
  );
}

Future<PreparedUpdate> downloadUpdate(
  UpdateCheckResult update, {
  void Function(int received, int? total)? onProgress,
}) async {
  if (!update.canInstall) {
    throw StateError('This release does not include a verified installer.');
  }
  final directory = await Directory.systemTemp.createTemp('wade-update-');
  try {
    final checksumFile = File(
      '${directory.path}${Platform.pathSeparator}checksum.sha256',
    );
    await _downloadFile(
      update.checksumUrl!,
      checksumFile,
      maximumBytes: 16 * 1024,
    );
    final expectedHash = parseSha256Checksum(
      await checksumFile.readAsString(),
      expectedFileName: update.installerFileName!,
    );
    final installer = File(
      '${directory.path}${Platform.pathSeparator}${update.installerFileName!}',
    );
    await _downloadFile(update.downloadUrl!, installer, onProgress: onProgress);
    final actualHash = await calculateSha256(installer);
    if (actualHash != expectedHash) {
      throw const FormatException(
        'The downloaded installer failed SHA-256 verification.',
      );
    }
    return PreparedUpdate(installer: installer, sha256: expectedHash);
  } on Object {
    if (directory.existsSync()) directory.deleteSync(recursive: true);
    rethrow;
  }
}

Future<void> launchUpdate(PreparedUpdate update) async {
  if (!Platform.isWindows) {
    throw UnsupportedError('Installing updates is supported on Windows.');
  }
  await Process.start(update.installer.path, const [
    '/SILENT',
    '/CLOSEAPPLICATIONS',
    '/RESTARTAPPLICATIONS',
    '/SUPPRESSMSGBOXES',
  ], mode: ProcessStartMode.detached);
}

String parseSha256Checksum(
  String contents, {
  required String expectedFileName,
}) {
  final pattern = RegExp(
    r'^\s*([0-9a-fA-F]{64})\s+[*]?(.+?)\s*$',
    multiLine: true,
  );
  for (final match in pattern.allMatches(contents)) {
    final fileName = match.group(2)!.trim();
    if (fileName.toLowerCase() == expectedFileName.toLowerCase()) {
      return match.group(1)!.toLowerCase();
    }
  }
  throw FormatException(
    'The release checksum does not contain $expectedFileName.',
  );
}

Future<String> calculateSha256(File file) async {
  if (!Platform.isWindows) {
    throw UnsupportedError('SHA-256 verification is supported on Windows.');
  }
  final result = await Process.run('certutil.exe', [
    '-hashfile',
    file.path,
    'SHA256',
  ]);
  if (result.exitCode != 0) {
    throw ProcessException(
      'certutil.exe',
      ['-hashfile', file.path, 'SHA256'],
      '${result.stderr}',
      result.exitCode,
    );
  }
  final match = RegExp(r'\b[0-9a-fA-F]{64}\b').firstMatch('${result.stdout}');
  if (match == null) {
    throw const FormatException('Windows returned an invalid SHA-256 hash.');
  }
  return match.group(0)!.toLowerCase();
}

Future<void> _downloadFile(
  String url,
  File destination, {
  int? maximumBytes,
  void Function(int received, int? total)? onProgress,
}) async {
  final uri = Uri.parse(url);
  if (uri.scheme != 'https') {
    throw FormatException('Update downloads must use HTTPS: $url');
  }
  final client = HttpClient();
  try {
    final request = await client.getUrl(uri);
    request.headers.set(
      HttpHeaders.userAgentHeader,
      'Windows-Audio-Display-Enforcer',
    );
    final response = await request.close();
    if (response.redirects.any(
      (redirect) => redirect.location.scheme != 'https',
    )) {
      throw const FormatException('Update download redirects must use HTTPS.');
    }
    if (response.statusCode != HttpStatus.ok) {
      throw HttpException(
        'Update download returned HTTP ${response.statusCode}.',
        uri: uri,
      );
    }
    final total = response.contentLength >= 0 ? response.contentLength : null;
    if (maximumBytes != null && total != null && total > maximumBytes) {
      throw const FormatException('The release checksum file is too large.');
    }
    final sink = destination.openWrite();
    var received = 0;
    try {
      await for (final chunk in response) {
        received += chunk.length;
        if (maximumBytes != null && received > maximumBytes) {
          throw const FormatException(
            'The release checksum file is too large.',
          );
        }
        sink.add(chunk);
        onProgress?.call(received, total);
      }
      await sink.flush();
    } finally {
      await sink.close();
    }
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
