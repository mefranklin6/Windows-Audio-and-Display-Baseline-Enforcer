import 'package:deployment_orchestrator_app/update_checker.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  test('requests an explicit relaunch from the silent update installer', () {
    expect(updateInstallerArguments, contains('/SILENT'));
    expect(updateInstallerArguments, contains('/RELAUNCHAPP=1'));
    expect(updateInstallerArguments, isNot(contains('/RESTARTAPPLICATIONS')));
  });

  test('compares release versions with optional v prefix and build number', () {
    expect(compareVersions('v1.2.0', '1.1.9'), greaterThan(0));
    expect(compareVersions('1.2.0', 'v1.2'), 0);
    expect(compareVersions('v2.0.0+4', '1.99.99'), greaterThan(0));
    expect(compareVersions('1.0.0', '1.0.1'), lessThan(0));
  });

  test('reports whether a newer release is available', () {
    const update = UpdateCheckResult(
      currentVersion: '1.0.0',
      latestVersion: 'v1.1.0',
      releaseUrl: releasesPageUrl,
      downloadUrl: null,
    );
    expect(update.updateAvailable, isTrue);
    expect(update.preferredUrl, releasesPageUrl);
    expect(update.canInstall, isFalse);
  });

  test('selects an installer and its matching checksum asset', () {
    final update = parseRelease({
      'tag_name': 'v1.2.0',
      'html_url': 'https://example.test/release',
      'assets': [
        {
          'name': 'Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator-1.2.0-Setup.exe.sha256',
          'browser_download_url': 'https://example.test/setup.sha256',
        },
        {
          'name': 'Windows-Audio-and-Display-Baseline-Enforcer-Orchestrator-1.2.0-Setup.exe',
          'browser_download_url': 'https://example.test/setup.exe',
        },
      ],
    });

    expect(update.downloadUrl, 'https://example.test/setup.exe');
    expect(update.checksumUrl, 'https://example.test/setup.sha256');
    expect(update.canInstall, isTrue);
  });

  test(
    'rejects an installer asset name that could escape the update folder',
    () {
      expect(
        () => parseRelease({
          'tag_name': 'v1.2.0',
          'assets': [
            {
              'name': r'..\app-1.2.0-Setup.exe',
              'browser_download_url': 'https://example.test/setup.exe',
            },
          ],
        }),
        throwsFormatException,
      );
    },
  );

  test('parses only the checksum for the expected installer', () {
    const expected =
        '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef';
    expect(
      parseSha256Checksum(
        '$expected  app-1.2.0-Setup.exe\n'
        'ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff  another.exe\n',
        expectedFileName: 'app-1.2.0-Setup.exe',
      ),
      expected,
    );
    expect(
      () => parseSha256Checksum(
        '$expected  another.exe',
        expectedFileName: 'app-1.2.0-Setup.exe',
      ),
      throwsFormatException,
    );
  });
}
