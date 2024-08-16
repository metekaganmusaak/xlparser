import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';

const String SHEET_NAME = 'Sheet1';

class HomeView extends StatefulWidget {
  const HomeView({super.key});

  @override
  State<HomeView> createState() => _HomeViewState();
}

class _HomeViewState extends State<HomeView> {
  // List<String> extracted = [];

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('XLParser'),
        centerTitle: true,
        actions: const [
          // IconButton(
          //   onPressed: () {
          //     setState(() {
          //       extracted.clear();
          //     });
          //   },
          //   icon: const Icon(Icons.clear),
          // ),
        ],
      ),

      body: Padding(
        padding: EdgeInsets.symmetric(
          horizontal: MediaQuery.sizeOf(context).width * 0.2,
        ),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          crossAxisAlignment: CrossAxisAlignment.center,
          children: [
            Align(
              child: FilledButton.icon(
                onPressed: () async {
                  final pickedFiles = await FilePicker.platform.pickFiles(
                    type: FileType.custom,
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
                style: ButtonStyle(
                    fixedSize: WidgetStateProperty.all(const Size(300, 50)),
                    backgroundColor: WidgetStateProperty.resolveWith((state) {
                      if (state.contains(WidgetState.hovered)) {
                        return Colors.green[800];
                      } else {
                        return Colors.green;
                      }
                    })),
                icon: const Icon(Icons.file_upload),
                label: const Text('Select Excel Files'),
              ),
            ),
          ],
        ),
      ),
      // body: SingleChildScrollView(
      //   padding: const EdgeInsets.all(16),
      //   child: Column(
      //     children: [
      //       Text(extracted.join('\n')),
      //     ],
      //   ),
      // ),
    );
  }

  void parseExcel(PlatformFile file) async {
    Uint8List? fileBytes;

    Excel? excel;

    if (kIsWeb) {
      fileBytes = file.bytes;

      if (fileBytes == null) {
        return;
      }
    }

    if (!kIsWeb) {
      final filePath = file.path;

      if (filePath == null) {
        return;
      }

      fileBytes = await File(filePath).readAsBytes();
    }

    if (fileBytes == null) {
      return;
    }

    excel = Excel.decodeBytes(fileBytes);

    int index = 1;

    final newExcel = Excel.createExcel();
    newExcel.setDefaultSheet(SHEET_NAME);

    CellStyle cellStyle = CellStyle(
      bold: true,
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 12,
      backgroundColorHex: ExcelColor.amber600,
      fontColorHex: ExcelColor.black,
    );

    newExcel.updateCell(
      SHEET_NAME,
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      const TextCellValue('TARİH'),
      cellStyle: cellStyle,
    );

    newExcel.updateCell(
      SHEET_NAME,
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 0),
      const TextCellValue('PARA'),
      cellStyle: cellStyle,
    );

    newExcel.updateCell(
      SHEET_NAME,
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
      const TextCellValue('FİRMA'),
      cellStyle: cellStyle,
    );

    for (var table in excel.tables.keys) {
      print('Max columns: ${excel.tables[table]!.maxColumns}'); // 17
      print('Max rows: ${excel.tables[table]!.maxRows}'); // 36

      for (var i = 7; i <= excel.tables[table]!.maxRows; i++) {
        final row = excel.tables[table]!.rows[i];

        //4th column
        final dateCell = row[3]?.value as TextCellValue?;
        // 7th column
        final moneyCell = row[6]?.value as IntCellValue?;
        // 16th column
        final companyCell = row[16]?.value as TextCellValue?;

        if (dateCell == null || moneyCell == null || companyCell == null) {
          break;
        }

        String date = dateCell.toString() == "null" ? '' : dateCell.toString();
        String money =
            moneyCell.toString() == "null" ? '' : moneyCell.toString();
        String company =
            companyCell.toString() == "null" ? '' : companyCell.toString();

        if (date.isEmpty || money.isEmpty || company.isEmpty) {
          break;
        }

        company = extractCompanyNameFromText(company);

        newExcel.insertRowIterables(
          SHEET_NAME,
          [
            TextCellValue(date),
            IntCellValue(int.tryParse(money) ?? 0),
            TextCellValue(company),
          ],
          index,
        );

        index++;
      }
    }

    newExcel.save(fileName: 'output.xlsx');
  }

  String extractCompanyNameFromText(String text) {
    final expression1 = RegExp(r'sorgu numaralı (.*?) tarafından');
    final expression2 = RegExp(r'(\d+) Nolu Müşterinin');
    final expression3 = RegExp(r'nolu (.*?) hesabından');

    if (expression1.hasMatch(text)) {
      final match = expression1.firstMatch(text);
      return match?.group(1) ?? text;
    }

    if (expression2.hasMatch(text)) {
      final match = expression2.firstMatch(text);
      return match?.group(0) ?? text;
    }

    if (expression3.hasMatch(text)) {
      final match = expression3.firstMatch(text);
      return match?.group(1) ?? text;
    }

    return text;
  }
}
