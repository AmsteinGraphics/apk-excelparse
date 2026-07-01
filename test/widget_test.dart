import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:apk_excelparse/main.dart';

void main() {
  testWidgets('Settings page renders', (WidgetTester tester) async {
    await tester.pumpWidget(const ApkExcelParseApp());

    expect(find.text('Settings'), findsOneWidget);
  });
}
