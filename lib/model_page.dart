import 'dart:io';
import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:fuzzywuzzy/fuzzywuzzy.dart';

class MyModelPage extends StatefulWidget {
  const MyModelPage({super.key});

  @override
  _MyModelPageState createState() => _MyModelPageState();
}

class _MyModelPageState extends State<MyModelPage> {
  bool _isProcessing = false;
  String _statusMessage = "";
  List<Map<String, dynamic>> uploadedFiles = [];
  TextCellValue(String value) {
    return Text(
        value); // Example of returning a Text widget with the given value
  }

  /// 📌 Chọn file Excel
  void _pickExcelFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );

    if (result != null) {
      setState(() {
        uploadedFiles = [
          {
            'name': result.files.single.name,
            'bytes': base64Encode(result
                .files.single.bytes!), // Chuyển đổi bytes thành chuỗi base64
          }
        ]; // Luôn chỉ chứa 1 file
      });

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Đã chọn file: ${result.files.single.name}')),
      );
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Không có file nào được chọn')),
      );
    }
  }

  /// 📌 Đọc dữ liệu từ file Excel
  Future<List<List<String>>> readExcelFromFile(List<int> fileBytes) async {
    var excel = Excel.decodeBytes(fileBytes);

    List<String> inputProducts = [];
    List<String> outputProducts = [];

    for (var table in excel.tables.keys) {
      var sheet = excel.tables[table]!;
      for (var row in sheet.rows) {
        if (row.isNotEmpty) {
          inputProducts.add(row[0]?.value.toString() ?? '');
          if (row.length > 1) {
            outputProducts.add(row[1]?.value.toString() ?? '');
          }
        }
      }
    }
    return [inputProducts, outputProducts];
  }

  /// 📌 Nhóm sản phẩm đầu ra theo sản phẩm đầu vào gần nhất
  Map<String, List<String>> groupOutputProducts(
      List<String> inputProducts, List<String> outputProducts) {
    Map<String, List<String>> groupedProducts = {};

    for (var product in inputProducts) {
      groupedProducts[product] = [];
    }

    List<String> ungroupedProducts = [];

    for (var output in outputProducts) {
      String bestMatch = "";
      int bestSimilarity = 0;

      for (var input in inputProducts) {
        int similarity = ratio(input, output);
        if (similarity > bestSimilarity) {
          bestSimilarity = similarity;
          bestMatch = input;
        }
      }

      if (bestSimilarity >= 80) {
        groupedProducts[bestMatch]!.add(output);
      } else {
        ungroupedProducts.add(output);
      }
    }

    // Thêm các sản phẩm không trùng vào nhóm riêng
    groupedProducts["Sản phẩm chưa phân loại"] = ungroupedProducts;
    return groupedProducts;
  }

  /// 📌 Xuất danh sách ra file Excel mới
  void writeExcel(String filePath, Map<String, List<String>> groupedProducts) {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    for (var entry in groupedProducts.entries) {
      sheet.appendRow(
          [TextCellValue(entry.key), TextCellValue(entry.value.join(", "))]);
    }

    var file = File(filePath);
    file.writeAsBytesSync(excel.encode()!);
  }

  /// 📌 Hàm xử lý quá trình đọc, nhóm và ghi file Excel
  Future<void> processExcel() async {
    if (uploadedFiles.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Vui lòng chọn một file Excel')),
      );
      return;
    }

    setState(() {
      _isProcessing = true;
      _statusMessage = "Đang xử lý...";
    });

    String outputFilePath =
        "assets/grouped_products.xlsx"; // Đường dẫn file Excel đầu ra

    try {
      List<int> fileBytes =
          base64Decode(uploadedFiles[0]['bytes']); // Chuyển đổi base64 về bytes
      List<List<String>> data = await readExcelFromFile(fileBytes);
      List<String> inputProducts = data[0];
      List<String> outputProducts = data[1];

      Map<String, List<String>> groupedProducts =
          groupOutputProducts(inputProducts, outputProducts);
      writeExcel(outputFilePath, groupedProducts);

      setState(() {
        _statusMessage = "✅ Hoàn tất! Kết quả được lưu tại: $outputFilePath";
      });
    } catch (e) {
      setState(() {
        _statusMessage = "❌ Lỗi: $e";
      });
    } finally {
      setState(() {
        _isProcessing = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text('TFLite Model Example')),
      body: Center(
        child: _isProcessing
            ? CircularProgressIndicator()
            : Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  ElevatedButton(
                    onPressed: _pickExcelFile,
                    child: Text("Chọn file Excel"),
                  ),
                  ElevatedButton(
                    onPressed: processExcel,
                    child: Text("Xử lý file Excel"),
                  ),
                  SizedBox(height: 20),
                  Text(
                    _statusMessage,
                    style: TextStyle(fontSize: 16, color: Colors.blue),
                  ),
                ],
              ),
      ),
    );
  }
}
