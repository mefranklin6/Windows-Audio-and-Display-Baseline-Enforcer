import 'dart:math' as math;

import 'package:flutter/material.dart';

class DisplayConfiguration {
  const DisplayConfiguration({
    required this.mode,
    required this.profilePath,
    required this.monitors,
  });

  factory DisplayConfiguration.fromJson(Map<String, dynamic> json) {
    final rawMonitors = json['monitors'];
    final monitorItems = rawMonitors is List
        ? rawMonitors
        : rawMonitors is Map
        ? [rawMonitors]
        : const <dynamic>[];
    return DisplayConfiguration(
      mode: json['mode'] as String? ?? 'Unknown display mode',
      profilePath: json['profile_path'] as String? ?? '',
      monitors:
          monitorItems
              .whereType<Map>()
              .map(
                (item) =>
                    DisplayMonitor.fromJson(Map<String, dynamic>.from(item)),
              )
              .toList()
            ..sort((a, b) => a.number.compareTo(b.number)),
    );
  }

  final String mode;
  final String profilePath;
  final List<DisplayMonitor> monitors;
}

class DisplayMonitor {
  const DisplayMonitor({
    required this.number,
    required this.name,
    required this.width,
    required this.height,
    required this.x,
    required this.y,
    required this.rotation,
    required this.refreshRate,
    required this.primary,
  });

  factory DisplayMonitor.fromJson(Map<String, dynamic> json) {
    int integer(String key) => (json[key] as num?)?.toInt() ?? 0;

    return DisplayMonitor(
      number: integer('number'),
      name: json['name'] as String? ?? '',
      width: integer('width'),
      height: integer('height'),
      x: integer('x'),
      y: integer('y'),
      rotation: integer('rotation'),
      refreshRate: (json['refresh_rate'] as num?)?.toDouble(),
      primary: json['primary'] as bool? ?? false,
    );
  }

  final int number;
  final String name;
  final int width;
  final int height;
  final int x;
  final int y;
  final int rotation;
  final double? refreshRate;
  final bool primary;

  String get label => name.isEmpty ? 'Display $number' : name;
  String get resolution => '$width × $height';
}

class DisplayConfigurationLoadResult {
  const DisplayConfigurationLoadResult({
    required this.configuration,
    this.error = '',
  });

  final DisplayConfiguration? configuration;
  final String error;
}

typedef DisplayConfigurationLoader =
    Future<DisplayConfigurationLoadResult> Function();

class DisplayConfigurationPage extends StatefulWidget {
  const DisplayConfigurationPage({
    required this.pc,
    required this.configuration,
    this.error = '',
    this.loader,
    this.asDialog = false,
    super.key,
  });

  final String pc;
  final DisplayConfiguration? configuration;
  final String error;
  final DisplayConfigurationLoader? loader;
  final bool asDialog;

  @override
  State<DisplayConfigurationPage> createState() =>
      _DisplayConfigurationPageState();
}

class _DisplayConfigurationPageState extends State<DisplayConfigurationPage> {
  late DisplayConfiguration? _configuration = widget.configuration;
  late String _error = widget.error;
  bool _loading = false;

  @override
  void initState() {
    super.initState();
    if (_configuration == null && _error.isEmpty && widget.loader != null) {
      _loading = true;
      _load();
    }
  }

  Future<void> _load() async {
    try {
      final result = await widget.loader!();
      if (!mounted) return;
      setState(() {
        _configuration = result.configuration;
        _error = result.error;
        _loading = false;
      });
    } on Object catch (error) {
      if (!mounted) return;
      setState(() {
        _error = '$error';
        _loading = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    final config = _configuration;
    final content = _buildContent(context, config);
    if (widget.asDialog) {
      return AlertDialog(
        title: Row(
          children: [
            const Icon(Icons.desktop_windows_outlined),
            const SizedBox(width: 10),
            Expanded(child: Text('Display configuration · ${widget.pc}')),
          ],
        ),
        content: SizedBox(width: 920, height: 720, child: content),
        actions: [
          FilledButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('Close'),
          ),
        ],
      );
    }
    return Scaffold(
      appBar: AppBar(title: Text('${widget.pc} display configuration')),
      body: SafeArea(child: content),
    );
  }

  Widget _buildContent(BuildContext context, DisplayConfiguration? config) {
    if (_loading) {
      return const Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            CircularProgressIndicator(),
            SizedBox(height: 16),
            Text('Reading saved display profile…'),
          ],
        ),
      );
    }
    return SingleChildScrollView(
      padding: EdgeInsets.all(widget.asDialog ? 8 : 24),
      child: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 1000),
          child: config == null
              ? _DisplayLoadError(error: _error)
              : Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    _ModeSummary(configuration: config),
                    const SizedBox(height: 20),
                    Text(
                      'Display arrangement',
                      style: Theme.of(context).textTheme.titleLarge,
                    ),
                    const SizedBox(height: 8),
                    if (config.monitors.isEmpty)
                      const _DisplayLoadError(
                        error: 'The saved profile does not contain an active display.',
                      )
                    else
                      MonitorArrangement(monitors: config.monitors),
                    if (config.monitors.isNotEmpty) ...[
                      const SizedBox(height: 20),
                      Text(
                        'Displays',
                        style: Theme.of(context).textTheme.titleLarge,
                      ),
                      const SizedBox(height: 8),
                      Wrap(
                        spacing: 12,
                        runSpacing: 12,
                        children: config.monitors
                            .map((monitor) => _MonitorDetails(monitor: monitor))
                            .toList(),
                      ),
                    ],
                  ],
                ),
        ),
      ),
    );
  }
}

class _ModeSummary extends StatelessWidget {
  const _ModeSummary({required this.configuration});

  final DisplayConfiguration configuration;

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(18),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            CircleAvatar(
              backgroundColor: scheme.primaryContainer,
              foregroundColor: scheme.onPrimaryContainer,
              child: const Icon(Icons.desktop_windows_outlined),
            ),
            const SizedBox(width: 14),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    configuration.mode,
                    key: const Key('displayConfigurationMode'),
                    style: Theme.of(context).textTheme.titleMedium
                        ?.copyWith(fontWeight: FontWeight.w700),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    '${configuration.monitors.length} active display${configuration.monitors.length == 1 ? '' : 's'}',
                  ),
                  if (configuration.profilePath.isNotEmpty) ...[
                    const SizedBox(height: 6),
                    Text(
                      configuration.profilePath,
                      style: Theme.of(context).textTheme.bodySmall,
                    ),
                  ],
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}

class MonitorArrangement extends StatelessWidget {
  const MonitorArrangement({required this.monitors, super.key});

  final List<DisplayMonitor> monitors;

  @override
  Widget build(BuildContext context) {
    final groups = _MonitorGroup.fromMonitors(monitors);
    final minX = groups.map((group) => group.x).reduce(math.min);
    final minY = groups.map((group) => group.y).reduce(math.min);
    final maxX = groups.map((group) => group.x + group.width).reduce(math.max);
    final maxY = groups.map((group) => group.y + group.height).reduce(math.max);
    final virtualWidth = math.max(1, maxX - minX);
    final virtualHeight = math.max(1, maxY - minY);

    return Semantics(
      label: monitors
          .map(
            (monitor) =>
                'Display ${monitor.number}, ${monitor.resolution}, position ${monitor.x}, ${monitor.y}${monitor.primary ? ', primary' : ''}',
          )
          .join('. '),
      child: Card(
        child: SizedBox(
          height: 360,
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: LayoutBuilder(
              builder: (context, constraints) {
                final scale = math.min(
                  constraints.maxWidth / virtualWidth,
                  constraints.maxHeight / virtualHeight,
                );
                final drawingWidth = virtualWidth * scale;
                final drawingHeight = virtualHeight * scale;
                final leftInset = (constraints.maxWidth - drawingWidth) / 2;
                final topInset = (constraints.maxHeight - drawingHeight) / 2;
                return Stack(
                  key: const Key('monitorArrangement'),
                  children: groups.map((group) {
                    return Positioned(
                      left: leftInset + (group.x - minX) * scale,
                      top: topInset + (group.y - minY) * scale,
                      width: group.width * scale,
                      height: group.height * scale,
                      child: _MonitorShape(group: group),
                    );
                  }).toList(),
                );
              },
            ),
          ),
        ),
      ),
    );
  }
}

class _MonitorShape extends StatelessWidget {
  const _MonitorShape({required this.group});

  final _MonitorGroup group;

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    final primary = group.monitors.any((monitor) => monitor.primary);
    final numbers = group.monitors.map((monitor) => monitor.number).join(' | ');
    return Tooltip(
      message: group.monitors.map((monitor) => monitor.label).join(' · '),
      child: Container(
        decoration: BoxDecoration(
          color: scheme.primaryContainer,
          borderRadius: BorderRadius.circular(7),
          border: Border.all(
            color: primary ? scheme.primary : scheme.outline,
            width: primary ? 3 : 1.5,
          ),
          boxShadow: [
            BoxShadow(
              color: scheme.shadow.withValues(alpha: 0.18),
              blurRadius: 5,
              offset: const Offset(0, 2),
            ),
          ],
        ),
        padding: const EdgeInsets.all(5),
        child: FittedBox(
          fit: BoxFit.scaleDown,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text(
                numbers,
                style: TextStyle(
                  color: scheme.onPrimaryContainer,
                  fontSize: 30,
                  fontWeight: FontWeight.w700,
                ),
              ),
              Text(
                '${group.width} × ${group.height}',
                style: TextStyle(color: scheme.onPrimaryContainer),
              ),
              if (primary)
                Text(
                  'Main display',
                  style: TextStyle(
                    color: scheme.onPrimaryContainer,
                    fontWeight: FontWeight.w600,
                  ),
                ),
            ],
          ),
        ),
      ),
    );
  }
}

class _MonitorDetails extends StatelessWidget {
  const _MonitorDetails({required this.monitor});

  final DisplayMonitor monitor;

  @override
  Widget build(BuildContext context) {
    final refreshRate = monitor.refreshRate;
    final refreshText = refreshRate == null
        ? 'Unknown refresh rate'
        : '${refreshRate.toStringAsFixed(refreshRate == refreshRate.roundToDouble() ? 0 : 2)} Hz';
    return SizedBox(
      width: 300,
      child: Card(
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              CircleAvatar(child: Text('${monitor.number}')),
              const SizedBox(width: 12),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      monitor.label,
                      style: const TextStyle(fontWeight: FontWeight.w700),
                    ),
                    Text('${monitor.resolution} · $refreshText'),
                    Text('Position ${monitor.x}, ${monitor.y}'),
                    if (monitor.rotation != 0)
                      Text('Rotated ${monitor.rotation}°'),
                    if (monitor.primary)
                      Text(
                        'Main display',
                        style: TextStyle(
                          color: Theme.of(context).colorScheme.primary,
                          fontWeight: FontWeight.w600,
                        ),
                      ),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _DisplayLoadError extends StatelessWidget {
  const _DisplayLoadError({required this.error});

  final String error;

  @override
  Widget build(BuildContext context) {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Icon(
              Icons.error_outline,
              color: Theme.of(context).colorScheme.error,
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Text(
                error.isEmpty
                    ? 'The saved display profile could not be read.'
                    : 'The saved display profile could not be read: $error',
              ),
            ),
          ],
        ),
      ),
    );
  }
}

class _MonitorGroup {
  const _MonitorGroup({
    required this.x,
    required this.y,
    required this.width,
    required this.height,
    required this.monitors,
  });

  factory _MonitorGroup.first(DisplayMonitor monitor) => _MonitorGroup(
    x: monitor.x,
    y: monitor.y,
    width: monitor.width,
    height: monitor.height,
    monitors: [monitor],
  );

  final int x;
  final int y;
  final int width;
  final int height;
  final List<DisplayMonitor> monitors;

  static List<_MonitorGroup> fromMonitors(List<DisplayMonitor> monitors) {
    final groups = <String, _MonitorGroup>{};
    for (final monitor in monitors) {
      final key =
          '${monitor.x}:${monitor.y}:${monitor.width}:${monitor.height}';
      final group = groups[key];
      if (group == null) {
        groups[key] = _MonitorGroup.first(monitor);
      } else {
        group.monitors.add(monitor);
      }
    }
    return groups.values.toList();
  }
}
