import 'dart:convert';

import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'package:url_launcher/url_launcher.dart';

void main() {
  runApp(const ApkExcelParseApp());
}

class ApkExcelParseApp extends StatelessWidget {
  const ApkExcelParseApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'APK Excel Parse',
      theme: ThemeData(useMaterial3: true, colorSchemeSeed: Colors.blue),
      home: const SettingsPage(),
    );
  }
}

class SettingsPage extends StatefulWidget {
  const SettingsPage({super.key});

  @override
  State<SettingsPage> createState() => _SettingsPageState();
}

class _SettingsPageState extends State<SettingsPage> {
  static const String _currentVersion = '1.0.0';
  bool _isChecking = false;

  Future<void> _checkForUpdates() async {
    setState(() => _isChecking = true);

    try {
      final release = await GitHubReleaseService.fetchLatestRelease();
      if (!mounted) {
        return;
      }

      if (release == null) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('No release information was available from GitHub.')),
        );
        return;
      }

      final latestVersion = release.tagName.replaceFirst(RegExp(r'^v'), '');
      final currentVersion = _currentVersion.replaceFirst(RegExp(r'^v'), '');
      final hasUpdate = latestVersion != currentVersion;

      if (hasUpdate) {
        await launchUrl(
          Uri.parse(release.htmlUrl),
          mode: LaunchMode.externalApplication,
        );
        if (!mounted) {
          return;
        }
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('Update available: ${release.tagName}. The release page has been opened.'),
          ),
        );
      } else {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('This app is already up to date.')),
        );
      }
    } catch (error) {
      if (!mounted) {
        return;
      }
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Unable to check for updates: $error')),
      );
    } finally {
      if (mounted) {
        setState(() => _isChecking = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Settings')),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          Card(
            child: Padding(
              padding: const EdgeInsets.all(16),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'App updates',
                    style: Theme.of(context).textTheme.titleMedium,
                  ),
                  const SizedBox(height: 8),
                  Text('Installed version: $_currentVersion'),
                  const SizedBox(height: 12),
                  FilledButton.icon(
                    onPressed: _isChecking ? null : _checkForUpdates,
                    icon: _isChecking
                        ? const SizedBox(
                            width: 16,
                            height: 16,
                            child: CircularProgressIndicator(strokeWidth: 2),
                          )
                        : const Icon(Icons.system_update_alt),
                    label: Text(_isChecking ? 'Checking…' : 'Update from GitHub'),
                  ),
                  const SizedBox(height: 8),
                  Text(
                    'This option checks the latest release from the GitHub repository and opens the release page for installation.',
                    style: Theme.of(context).textTheme.bodySmall,
                  ),
                ],
              ),
            ),
          ),
          const SizedBox(height: 16),
          SwitchListTile.adaptive(
            value: true,
            onChanged: (_) {},
            title: const Text('Check for updates automatically'),
            subtitle: const Text('Enable periodic checks when the app launches.'),
          ),
        ],
      ),
    );
  }
}

class GitHubReleaseService {
  static Future<ReleaseInfo?> fetchLatestRelease() async {
    final response = await http.get(
      Uri.parse('https://api.github.com/repos/emy/apk-excelparse/releases/latest'),
    );

    if (response.statusCode != 200) {
      return null;
    }

    final body = jsonDecode(response.body) as Map<String, dynamic>;
    return ReleaseInfo(
      tagName: body['tag_name']?.toString() ?? 'unknown',
      name: body['name']?.toString() ?? 'Latest release',
      htmlUrl: body['html_url']?.toString() ?? 'https://github.com/emy/apk-excelparse/releases',
    );
  }
}

class ReleaseInfo {
  const ReleaseInfo({required this.tagName, required this.name, required this.htmlUrl});

  final String tagName;
  final String name;
  final String htmlUrl;
}
