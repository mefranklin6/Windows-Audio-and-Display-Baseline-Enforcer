import 'package:deployment_orchestrator_app/update_checker.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
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
  });
}
