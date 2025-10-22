import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'dart:io';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Reader App',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
      ),
      home: const ExcelReaderPage(),
    );
  }
}

class ExcelReaderPage extends StatefulWidget {
  const ExcelReaderPage({super.key});

  @override
  State<ExcelReaderPage> createState() => _ExcelReaderPageState();
}

class _ExcelReaderPageState extends State<ExcelReaderPage> {
  List<List<dynamic>> _excelData = [];

  Future<void> _pickAndReadExcel() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null) {
      File file = File(result.files.single.path!);
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      List<List<dynamic>> data = [];
      for (var table in excel.tables.keys) {
        for (var row in excel.tables[table]!.rows) {
          data.add(row.map((cell) => cell?.value ?? '').toList());
        }
      }

      setState(() {
        _excelData = data;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel Reader'),
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
      ),
      body: _excelData.isEmpty
          ? Center(
              child: ElevatedButton(
                onPressed: _pickAndReadExcel,
                child: const Text('Pick Excel File'),
              ),
            )
          : ListView.builder(
              itemCount: _excelData.length,
              itemBuilder: (context, index) {
                return Card(
                  margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                  child: Padding(
                    padding: const EdgeInsets.all(16),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          'Row ${index + 1}',
                          style: Theme.of(context).textTheme.titleMedium,
                        ),
                        const SizedBox(height: 8),
                        ..._excelData[index].map((cell) => Text(
                              '$cell',
                              style: Theme.of(context).textTheme.bodyMedium,
                            )),
                      ],
                    ),
                  ),
                );
              },
            ),
    );
  }
}