import 'dart:convert';
import 'dart:io';

import 'package:flutter/material.dart';

import 'log_formatting.dart';

enum AudioTextEncoding { utf8, utf8Bom, utf16LittleEndian, utf16BigEndian }

class AudioTextFile {
  const AudioTextFile(this.text, this.encoding);

  final String text;
  final AudioTextEncoding encoding;
}

AudioTextFile decodeAudioText(List<int> bytes) {
  if (bytes.length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe) {
    return AudioTextFile(
      _decodeUtf16(bytes, offset: 2, littleEndian: true),
      AudioTextEncoding.utf16LittleEndian,
    );
  }
  if (bytes.length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff) {
    return AudioTextFile(
      _decodeUtf16(bytes, offset: 2, littleEndian: false),
      AudioTextEncoding.utf16BigEndian,
    );
  }
  if (bytes.length >= 3 &&
      bytes[0] == 0xef &&
      bytes[1] == 0xbb &&
      bytes[2] == 0xbf) {
    return AudioTextFile(
      utf8.decode(bytes.sublist(3)),
      AudioTextEncoding.utf8Bom,
    );
  }
  // Handle BOM-less UTF-16 JSON as well. JSON starts with an ASCII object or
  // array delimiter, sometimes preceded by ASCII whitespace.
  if (bytes.length >= 4 && bytes[1] == 0 && bytes[3] == 0) {
    return AudioTextFile(
      _decodeUtf16(bytes, offset: 0, littleEndian: true),
      AudioTextEncoding.utf16LittleEndian,
    );
  }
  if (bytes.length >= 4 && bytes[0] == 0 && bytes[2] == 0) {
    return AudioTextFile(
      _decodeUtf16(bytes, offset: 0, littleEndian: false),
      AudioTextEncoding.utf16BigEndian,
    );
  }
  return AudioTextFile(utf8.decode(bytes), AudioTextEncoding.utf8);
}

String _decodeUtf16(
  List<int> bytes, {
  required int offset,
  required bool littleEndian,
}) {
  if ((bytes.length - offset).isOdd) {
    throw const FormatException(
      'UTF-16 file contains an incomplete code unit.',
    );
  }
  final codeUnits = <int>[];
  for (var index = offset; index < bytes.length; index += 2) {
    codeUnits.add(
      littleEndian
          ? bytes[index] | (bytes[index + 1] << 8)
          : (bytes[index] << 8) | bytes[index + 1],
    );
  }
  return String.fromCharCodes(codeUnits);
}

List<int> encodeAudioText(String text, AudioTextEncoding encoding) {
  if (encoding == AudioTextEncoding.utf8) return utf8.encode(text);
  if (encoding == AudioTextEncoding.utf8Bom) {
    return <int>[0xef, 0xbb, 0xbf, ...utf8.encode(text)];
  }
  final littleEndian = encoding == AudioTextEncoding.utf16LittleEndian;
  final bytes = <int>[
    if (littleEndian) ...[0xff, 0xfe] else ...[0xfe, 0xff],
  ];
  for (final codeUnit in text.codeUnits) {
    if (littleEndian) {
      bytes.addAll([codeUnit & 0xff, codeUnit >> 8]);
    } else {
      bytes.addAll([codeUnit >> 8, codeUnit & 0xff]);
    }
  }
  return bytes;
}

const audioLevelNames = <String, String>{
  'PlaybackVolume': 'Playback',
  'PlaybackCommunicationVolume': 'Playback communications',
  'RecordingVolume': 'Recording',
  'RecordingCommunicationVolume': 'Recording communications',
};

const audioMuteKeys = <String, String>{
  'PlaybackVolume': 'PlaybackMute',
  'PlaybackCommunicationVolume': 'PlaybackCommunicationMute',
  'RecordingVolume': 'RecordingMute',
  'RecordingCommunicationVolume': 'RecordingCommunicationMute',
};

class AudioConfigurationData {
  AudioConfigurationData({
    required this.levels,
    required this.devices,
    required this.levelsJson,
    required this.devicesJson,
    required this.log,
  });

  factory AudioConfigurationData.fromJsonText({
    required String levelsJson,
    required String devicesJson,
    String? log,
  }) {
    final decodedLevels = jsonDecode(levelsJson);
    final decodedDevices = jsonDecode(devicesJson);
    if (decodedLevels is! Map) {
      throw const FormatException('audio_levels.json must contain an object.');
    }
    // PowerShell serializes a single pipeline result as an object, while
    // multiple results become an array. Treat either shape as a device list.
    final deviceList = switch (decodedDevices) {
      List<dynamic> devices => devices,
      Map _ => <dynamic>[decodedDevices],
      _ => null,
    };
    if (deviceList == null) {
      throw const FormatException(
        'audio_device_list.json must contain a device object or an array.',
      );
    }
    return AudioConfigurationData.fromMaps(
      levels: Map<String, dynamic>.from(decodedLevels),
      devices: deviceList,
      levelsJson: levelsJson,
      devicesJson: devicesJson,
      log: log,
    );
  }

  factory AudioConfigurationData.fromMaps({
    required Map<String, dynamic> levels,
    required List<dynamic> devices,
    String? levelsJson,
    String? devicesJson,
    String? log,
  }) {
    final normalizedDevices = devices.map((item) {
      if (item is! Map) {
        throw const FormatException(
          'Every audio device must be a JSON object.',
        );
      }
      final device = Map<String, dynamic>.from(item);
      if (device['Name'] is! String || device['Type'] is! String) {
        throw const FormatException(
          'Every audio device must have string Name and Type fields.',
        );
      }
      return device;
    }).toList();
    for (final volumeKey in audioLevelNames.keys) {
      parseAudioVolume(levels[volumeKey], key: volumeKey);
      final muteKey = audioMuteKeys[volumeKey]!;
      final muteValue = levels[muteKey];
      if (muteValue != null && muteValue is! bool) {
        throw FormatException('$muteKey must be true, false, null, or absent.');
      }
    }
    return AudioConfigurationData(
      levels: levels,
      devices: normalizedDevices,
      levelsJson:
          levelsJson ?? const JsonEncoder.withIndent('  ').convert(levels),
      devicesJson:
          devicesJson ??
          const JsonEncoder.withIndent('  ').convert(normalizedDevices),
      log: log,
    );
  }

  final Map<String, dynamic> levels;
  final List<Map<String, dynamic>> devices;
  final String levelsJson;
  final String devicesJson;
  final String? log;
}

double? parseAudioVolume(Object? raw, {required String key}) {
  if (_isUnavailableAudioVolume(raw)) return null;
  final value = switch (raw) {
    num number => number.toDouble(),
    String text => double.tryParse(text.replaceAll('%', '').trim()),
    _ => null,
  };
  if (value == null || !value.isFinite || value < 0 || value > 100) {
    throw FormatException('$key must be a number from 0 to 100.');
  }
  return value;
}

bool _isUnavailableAudioVolume(Object? raw) {
  if (raw is! List || !raw.contains(null)) return false;
  return raw.whereType<String>().any(
    (message) => message.trimLeft().startsWith('No value found for'),
  );
}

abstract interface class AudioConfigurationGateway {
  Future<AudioConfigurationData> load(String pc);

  Future<void> save(
    String pc, {
    required String levelsJson,
    required String devicesJson,
  });
}

class FileAudioConfigurationGateway implements AudioConfigurationGateway {
  const FileAudioConfigurationGateway();

  String _ctsPath(String pc) {
    final target = pc.trim();
    if (!RegExp(r'^[A-Za-z0-9_.-]+$').hasMatch(target)) {
      throw ArgumentError.value(pc, 'pc', 'contains unsupported characters');
    }
    final localNames = <String>{
      'localhost',
      '.',
      '127.0.0.1',
      Platform.localHostname.toLowerCase(),
    };
    if (localNames.contains(target.toLowerCase())) {
      return r'C:\ProgramData\CTS';
    }
    return r'\\' + target + r'\C$\ProgramData\CTS';
  }

  @override
  Future<AudioConfigurationData> load(String pc) async {
    final folder = _ctsPath(pc);
    final levelsFile = File(
      '$folder${Platform.pathSeparator}audio_levels.json',
    );
    final devicesFile = File(
      '$folder${Platform.pathSeparator}audio_device_list.json',
    );
    final logFile = File(
      '$folder${Platform.pathSeparator}AudioDeviceStartup.log',
    );
    final results = await Future.wait([
      _readTextFile(levelsFile),
      _readTextFile(devicesFile),
    ]);
    String? log;
    try {
      log = (await _readTextFile(logFile)).text;
    } on FileSystemException {
      log = null;
    } on FormatException {
      log = null;
    }
    return AudioConfigurationData.fromJsonText(
      levelsJson: results[0].text,
      devicesJson: results[1].text,
      log: log,
    );
  }

  @override
  Future<void> save(
    String pc, {
    required String levelsJson,
    required String devicesJson,
  }) async {
    // Validate both documents before changing either remote file.
    AudioConfigurationData.fromJsonText(
      levelsJson: levelsJson,
      devicesJson: devicesJson,
    );
    final folder = _ctsPath(pc);
    final levelsFile = File(
      '$folder${Platform.pathSeparator}audio_levels.json',
    );
    final devicesFile = File(
      '$folder${Platform.pathSeparator}audio_device_list.json',
    );
    final oldLevels = await levelsFile.readAsBytes();
    final oldDevices = await devicesFile.readAsBytes();
    final levelsEncoding = decodeAudioText(oldLevels).encoding;
    final devicesEncoding = decodeAudioText(oldDevices).encoding;
    await _replaceFile(levelsFile, encodeAudioText(levelsJson, levelsEncoding));
    try {
      await _replaceFile(
        devicesFile,
        encodeAudioText(devicesJson, devicesEncoding),
      );
    } on Object {
      await _replaceFile(levelsFile, oldLevels);
      rethrow;
    }
  }

  Future<AudioTextFile> _readTextFile(File file) async {
    return decodeAudioText(await file.readAsBytes());
  }

  Future<void> _replaceFile(File destination, List<int> contents) async {
    final temporary = File('${destination.path}.$pid.tmp');
    try {
      await temporary.writeAsBytes(contents, flush: true);
      await temporary.rename(destination.path);
    } finally {
      if (await temporary.exists()) await temporary.delete();
    }
  }
}

class AudioConfigurationDialog extends StatefulWidget {
  const AudioConfigurationDialog({
    required this.pc,
    required this.gateway,
    super.key,
  });

  final String pc;
  final AudioConfigurationGateway gateway;

  @override
  State<AudioConfigurationDialog> createState() =>
      _AudioConfigurationDialogState();
}

class _AudioConfigurationDialogState extends State<AudioConfigurationDialog> {
  final _levelsController = TextEditingController();
  final _devicesController = TextEditingController();
  AudioConfigurationData? _data;
  Object? _error;
  bool _loading = true;
  bool _saving = false;
  bool _editing = false;
  bool _showJson = false;
  late Map<String, double?> _volumes;
  late Map<String, bool?> _mutes;
  late Map<String, int?> _deviceSelections;

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    _levelsController.dispose();
    _devicesController.dispose();
    super.dispose();
  }

  Future<void> _load() async {
    setState(() {
      _loading = true;
      _error = null;
    });
    try {
      final data = await widget.gateway.load(widget.pc);
      if (!mounted) return;
      _setData(data);
      setState(() => _loading = false);
    } on Object catch (error) {
      if (!mounted) return;
      setState(() {
        _loading = false;
        _error = error;
      });
    }
  }

  void _setData(AudioConfigurationData data) {
    _data = data;
    _levelsController.text = const JsonEncoder.withIndent('  ')
        .convert(data.levels);
    _devicesController.text = const JsonEncoder.withIndent('  ')
        .convert(data.devices);
    _volumes = {
      for (final key in audioLevelNames.keys)
        key: parseAudioVolume(data.levels[key], key: key),
    };
    _mutes = {
      for (final key in audioLevelNames.keys)
        key: data.levels[audioMuteKeys[key]] as bool?,
    };
    _deviceSelections = {
      for (final type in const ['Playback', 'Recording'])
        '$type:Default': _selectedDevice(data.devices, type, 'Default'),
      for (final type in const ['Playback', 'Recording'])
        '$type:DefaultCommunication': _selectedDevice(
          data.devices,
          type,
          'DefaultCommunication',
        ),
    };
  }

  int? _selectedDevice(
    List<Map<String, dynamic>> devices,
    String type,
    String property,
  ) {
    for (var index = 0; index < devices.length; index++) {
      if (devices[index]['Type'] == type && devices[index][property] == true) {
        return index;
      }
    }
    return null;
  }

  void _cancelEditing() {
    _setData(_data!);
    setState(() {
      _editing = false;
      _showJson = false;
      _error = null;
    });
  }

  Future<void> _save() async {
    String levelsJson;
    String devicesJson;
    try {
      if (_showJson) {
        levelsJson = _levelsController.text;
        devicesJson = _devicesController.text;
        AudioConfigurationData.fromJsonText(
          levelsJson: levelsJson,
          devicesJson: devicesJson,
        );
      } else {
        (levelsJson, devicesJson) = _guiJson();
      }
      setState(() {
        _saving = true;
        _error = null;
      });
      await widget.gateway.save(
        widget.pc,
        levelsJson: levelsJson,
        devicesJson: devicesJson,
      );
      final refreshed = await widget.gateway.load(widget.pc);
      if (!mounted) return;
      _setData(refreshed);
      setState(() {
        _saving = false;
        _editing = false;
        _showJson = false;
      });
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Audio configuration saved to ${widget.pc}.')),
      );
    } on Object catch (error) {
      if (!mounted) return;
      setState(() {
        _saving = false;
        _error = error;
      });
    }
  }

  (String, String) _guiJson() {
    final levels = Map<String, dynamic>.from(_data!.levels);
    for (final key in audioLevelNames.keys) {
      final volume = _volumes[key];
      if (volume != null) levels[key] = '${volume.round()}%';
      final mute = _mutes[key];
      if (mute != null) levels[audioMuteKeys[key]!] = mute;
    }
    final devices = _data!.devices
        .map((device) => Map<String, dynamic>.from(device))
        .toList();
    for (final entry in _deviceSelections.entries) {
      final parts = entry.key.split(':');
      final type = parts[0];
      final property = parts[1];
      for (var index = 0; index < devices.length; index++) {
        if (devices[index]['Type'] == type) {
          devices[index][property] = index == entry.value;
        }
      }
    }
    return (
      const JsonEncoder.withIndent('  ').convert(levels),
      const JsonEncoder.withIndent('  ').convert(devices),
    );
  }

  void _toggleJsonMode() {
    if (!_showJson) {
      final (levelsJson, devicesJson) = _guiJson();
      _levelsController.text = levelsJson;
      _devicesController.text = devicesJson;
      setState(() {
        _showJson = true;
        _editing = true;
        _error = null;
      });
      return;
    }
    try {
      final parsed = AudioConfigurationData.fromJsonText(
        levelsJson: _levelsController.text,
        devicesJson: _devicesController.text,
        log: _data!.log,
      );
      _setData(parsed);
      setState(() {
        _showJson = false;
        _error = null;
      });
    } on Object catch (error) {
      setState(() => _error = error);
    }
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: Row(
        children: [
          const Icon(Icons.graphic_eq_rounded),
          const SizedBox(width: 10),
          Expanded(child: Text('Audio configuration · ${widget.pc}')),
        ],
      ),
      content: SizedBox(width: 760, height: 610, child: _buildContent()),
      actions: [
        if (_data != null && !_loading) ...[
          OutlinedButton.icon(
            key: const Key('audioLogButton'),
            onPressed: _data!.log == null ? null : _showLog,
            icon: const Icon(Icons.description_outlined),
            label: const Text('Audio Log'),
          ),
          OutlinedButton.icon(
            key: const Key('audioJsonButton'),
            onPressed: _saving ? null : _toggleJsonMode,
            icon: const Icon(Icons.data_object_rounded),
            label: Text(_showJson ? 'GUI' : 'JSON'),
          ),
          if (!_editing)
            FilledButton.tonalIcon(
              key: const Key('editAudioButton'),
              onPressed: () => setState(() => _editing = true),
              icon: const Icon(Icons.edit_outlined),
              label: const Text('Edit'),
            )
          else ...[
            TextButton(
              onPressed: _saving ? null : _cancelEditing,
              child: const Text('Cancel'),
            ),
            FilledButton.icon(
              key: const Key('saveAudioButton'),
              onPressed: _saving ? null : _save,
              icon: _saving
                  ? const SizedBox.square(
                      dimension: 16,
                      child: CircularProgressIndicator(strokeWidth: 2),
                    )
                  : const Icon(Icons.save_outlined),
              label: const Text('Save'),
            ),
          ],
        ],
        TextButton(
          onPressed: _saving ? null : () => Navigator.of(context).pop(),
          child: const Text('Close'),
        ),
      ],
    );
  }

  Widget _buildContent() {
    if (_loading) return const Center(child: CircularProgressIndicator());
    if (_data == null) {
      return Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            const Icon(Icons.error_outline, size: 40),
            const SizedBox(height: 12),
            Text('Could not load the audio configuration.\n${_errorText()}'),
            const SizedBox(height: 12),
            OutlinedButton.icon(
              onPressed: _load,
              icon: const Icon(Icons.refresh),
              label: const Text('Retry'),
            ),
          ],
        ),
      );
    }
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        if (_error != null) ...[
          MaterialBanner(
            content: Text(_errorText()),
            leading: const Icon(Icons.error_outline),
            actions: [
              TextButton(
                onPressed: () => setState(() => _error = null),
                child: const Text('Dismiss'),
              ),
            ],
          ),
          const SizedBox(height: 8),
        ],
        Expanded(child: _showJson ? _buildJsonEditors() : _buildGui()),
      ],
    );
  }

  String _errorText() =>
      _error.toString().replaceFirst('FormatException: ', '');

  Widget _buildGui() {
    return SingleChildScrollView(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          Text(
            'Volume and mute',
            style: Theme.of(context).textTheme.titleMedium,
          ),
          const SizedBox(height: 8),
          for (final entry in audioLevelNames.entries)
            _buildLevel(entry.key, entry.value),
          const SizedBox(height: 18),
          Text(
            'Default devices',
            style: Theme.of(context).textTheme.titleMedium,
          ),
          const SizedBox(height: 8),
          _buildDeviceDropdown('Playback', 'Default', 'Playback'),
          _buildDeviceDropdown(
            'Playback',
            'DefaultCommunication',
            'Playback communications',
          ),
          _buildDeviceDropdown('Recording', 'Default', 'Recording'),
          _buildDeviceDropdown(
            'Recording',
            'DefaultCommunication',
            'Recording communications',
          ),
        ],
      ),
    );
  }

  Widget _buildLevel(String key, String label) {
    final value = _volumes[key];
    final muted = _mutes[key];
    if (value == null) {
      return Row(
        children: [
          SizedBox(width: 205, child: Text(label)),
          const Expanded(
            child: Text('Unavailable (no device)', textAlign: TextAlign.center),
          ),
          const SizedBox(width: 118),
        ],
      );
    }
    return Row(
      children: [
        SizedBox(width: 205, child: Text(label)),
        Expanded(
          child: Slider(
            key: Key('audioVolume-$key'),
            value: value,
            min: 0,
            max: 100,
            divisions: 100,
            label: '${value.round()}%',
            onChanged: _editing && !_saving
                ? (next) => setState(() => _volumes[key] = next)
                : null,
          ),
        ),
        SizedBox(width: 48, child: Text('${value.round()}%')),
        if (muted == null)
          const SizedBox(
            width: 70,
            child: Text('Unknown', textAlign: TextAlign.center),
          ),
        IconButton(
          key: Key('audioMute-$key'),
          tooltip: muted == null
              ? 'Mute status unknown (not recorded)'
              : muted
              ? 'Muted'
              : 'Unmuted',
          onPressed: _editing && !_saving
              ? () => setState(() => _mutes[key] = !(muted ?? false))
              : null,
          icon: Icon(switch (muted) {
            true => Icons.volume_off_rounded,
            false => Icons.volume_up_rounded,
            null => Icons.help_outline_rounded,
          }),
        ),
      ],
    );
  }

  Widget _buildDeviceDropdown(String type, String property, String label) {
    final choices = <DropdownMenuItem<int>>[];
    for (var index = 0; index < _data!.devices.length; index++) {
      final device = _data!.devices[index];
      if (device['Type'] == type) {
        choices.add(
          DropdownMenuItem(value: index, child: Text(device['Name'] as String)),
        );
      }
    }
    final selectionKey = '$type:$property';
    return Padding(
      padding: const EdgeInsets.only(bottom: 12),
      child: DropdownButtonFormField<int>(
        key: Key('audioDevice-$selectionKey'),
        initialValue: _deviceSelections[selectionKey],
        decoration: InputDecoration(labelText: label),
        isExpanded: true,
        items: choices,
        onChanged: _editing && !_saving
            ? (value) => setState(() => _deviceSelections[selectionKey] = value)
            : null,
      ),
    );
  }

  Widget _buildJsonEditors() {
    return Column(
      children: [
        Expanded(
          child: TextField(
            key: const Key('audioLevelsJsonEditor'),
            controller: _levelsController,
            expands: true,
            maxLines: null,
            minLines: null,
            readOnly: _saving,
            style: const TextStyle(fontFamily: 'monospace'),
            decoration: const InputDecoration(
              labelText: 'audio_levels.json',
              alignLabelWithHint: true,
            ),
          ),
        ),
        const SizedBox(height: 12),
        Expanded(
          child: TextField(
            key: const Key('audioDevicesJsonEditor'),
            controller: _devicesController,
            expands: true,
            maxLines: null,
            minLines: null,
            readOnly: _saving,
            style: const TextStyle(fontFamily: 'monospace'),
            decoration: const InputDecoration(
              labelText: 'audio_device_list.json',
              alignLabelWithHint: true,
            ),
          ),
        ),
      ],
    );
  }

  Future<void> _showLog() {
    return showDialog<void>(
      context: context,
      builder: (context) => AlertDialog(
        title: Text('AudioDeviceStartup.log · ${widget.pc}'),
        content: SizedBox(
          width: 720,
          height: 500,
          child: SingleChildScrollView(
            child: SelectableText.rich(
              TextSpan(children: buildLogSeveritySpans(_data!.log!)),
              key: const Key('audioLogContents'),
              style: const TextStyle(fontFamily: 'monospace'),
            ),
          ),
        ),
        actions: [
          FilledButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('Close'),
          ),
        ],
      ),
    );
  }
}
