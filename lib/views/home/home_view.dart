import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

class HomeView extends StatefulWidget {
  const HomeView({super.key});

  @override
  State<HomeView> createState() => _HomeViewState();
}

class _HomeViewState extends State<HomeView> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('XLParser'),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: () async {
          final pickedFiles = await FilePicker.platform.pickFiles(
            allowMultiple: true,
            allowedExtensions: ['xlsx'],
          );

          if (pickedFiles == null) {
            return;
          }

          if (pickedFiles.files.isEmpty) {
            return;
          }

          final List<PlatformFile> excels = pickedFiles.files;

          for (var excel in excels) {
            parseExcel(excel);
          }
        },
        child: const Icon(Icons.add),
      ),
    );
  }

  void parseExcel(PlatformFile file) async {
    final filePath = file.path;

    String extracted = '';
    int index = 1;

    if (filePath == null) {
      return;
    }

    var bytes = File(filePath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    for (var table in excel.tables.keys) {
      print('Max columns: ${excel.tables[table]!.maxColumns}'); // 17
      print('Max rows: ${excel.tables[table]!.maxRows}'); // 36

      for (var i = 8; i < excel.tables[table]!.maxRows; i++) {
        final row = excel.tables[table]!.rows[i];

        // 7th column
        final moneyCell = row[7]?.value;
        // 16th column
        final companyCell = row[16]?.value;

        String money =
            moneyCell.toString() == "null" ? '' : moneyCell.toString();
        String company =
            companyCell.toString() == "null" ? '' : companyCell.toString();

        if (money.isEmpty ||
            money == 'null' ||
            company.isEmpty ||
            company == 'null') {
          break;
        }

        if (company.contains('Nolu Müşterinin')) {
          final match = RegExp(r'\d+').firstMatch(company);
          if (match != null) {
            company = match.group(0) ?? '';
          }
        }

        // Get all upper case words
        final matches = RegExp(r'[A-Z]+').allMatches(company);
        for (var match in matches) {
          company = company.replaceAll(match.group(0)!, '');
        }

        extracted += '$index. $money - $company\n';
        index++;
      }
    }

    print('Exported: $extracted');
    // final exported = await File('export-${file.name}.txt').writeAsString(extracted);
  }
}
