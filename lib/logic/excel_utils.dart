import 'dart:convert';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as excel;
import 'dart:html' as html;

class ExcelUtils {
  /// Tạo và tải xuống file Excel
  static void createExcelFile(List<List<String>> dataExcel, String fileName,
      bool createManySheets, bool isShopping) {
    // Tạo một workbook
    final excel.Workbook workbook = excel.Workbook();

    // Accessing worksheet via index.
    final excel.Worksheet sheet = workbook.worksheets[0];
    sheet.name = fileName; // Đổi tên sheet

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < dataExcel.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < dataExcel[rowIndex].length - 2;
          colIndex++) {
        final cell = sheet.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(dataExcel[rowIndex][colIndex]);

        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }
    // Thêm dữ liệu vào sheet hàng hoá
    List<List<String>> transformedData = [];
    Set<String> seenCodeMH = {}; // Set lưu các mã mặt hàng đã xuất hiện
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        // Lấy mã khách hàng và tên khách hàng
        String codeMH; // Mã mặt hàng
        String nameMH; // Tên mặt hàng
        String unit; // Đơn vị tính
        String price; // Đơn giá
        if (isShopping) {
          codeMH = row[9]; // Mã mặt hàng
          nameMH = row[10]; // Tên mặt hàng
          unit = row[11]; // Đơn vị tính
          price = row[13]; // Đơn giá
        } else {
          codeMH = row[11]; // Mã mặt hàng
          nameMH = row[12]; // Tên mặt hàng
          unit = row[13]; // Đơn vị tính
          price = row[15]; // Đơn giá
        }

        // Nếu không phải hàng header và codeMH đã tồn tại thì bỏ qua
        if (i != 0 && seenCodeMH.contains(codeMH)) {
          continue;
        }
        // Thêm codeMH vào set để kiểm tra cho các hàng sau
        seenCodeMH.add(codeMH);

        List<String> newRow;
        if (i == 0) {
          newRow = [
            codeMH,
            nameMH,
            "Loại quy cách",
            "Thông số đặc tả",
            unit,
            "Danh mục mặt hàng",
            "Hàng đóng kiện",
            "Quản lý số lượng",
            "Quy trình",
            price,
            "Giá mua Tình trạng thuế",
            "Giá bán",
            "Giá bán Tình trạng thuế"
          ];
        } else {
          newRow = [
            codeMH,
            nameMH,
            "",
            "",
            unit,
            "3",
            "",
            "",
            "",
            price,
            "",
            "",
            ""
          ];
        }

        transformedData.add(newRow);
      }
    }

    // Thêm dữ liệu vào sheet nhà cung cấp
    List<List<String>> customerData = [];
    Set<String> seenCodeCustomer = {}; // Set lưu các mã mặt hàng đã xuất hiện
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        // Lấy mã khách hàng và tên khách hàng
        String codeCustomer = row[2]; // Mã KH/ NCC
        String nameCustomer = row[3]; // Tên công ty
        String person = row[4]; // Người phụ trách
        String customerAddress; // Địa chỉ công ty
        String customerPhone; // Điện thoại
        if (isShopping) {
          customerAddress = row[22]; // Địa chỉ công ty
          customerPhone = row[23]; // Điện thoại
        } else {
          customerAddress = row[26]; // Địa chỉ công ty
          customerPhone = row[27]; // Điện thoại
        }

        // Nếu không phải hàng header và codeMH đã tồn tại thì bỏ qua
        if (i != 0 && seenCodeCustomer.contains(codeCustomer)) {
          continue;
        }
        // Thêm codeMH vào set để kiểm tra cho các hàng sau
        seenCodeCustomer.add(codeCustomer);

        List<String> newRow;
        if (i == 0) {
          newRow = [
            codeCustomer,
            nameCustomer,
            "Điện thoại",
            "Mã 1",
            "Địa chỉ công ty",
            "Email",
            person,
            "Mã số thuế"
          ];
        } else {
          newRow = [
            codeCustomer,
            nameCustomer,
            customerPhone,
            "",
            customerAddress,
            "",
            person,
            codeCustomer
          ];
        }

        customerData.add(newRow);
      }
    }

    if (createManySheets) {
      final excel.Worksheet sheet2 = workbook.worksheets.add();
      sheet2.name = "Mặt hàng"; // Đặt tên sheet phụ
      // Thêm dữ liệu vào sheet
      for (int rowIndex = 0; rowIndex < transformedData.length; rowIndex++) {
        for (int colIndex = 0;
            colIndex < transformedData[rowIndex].length;
            colIndex++) {
          final cell = sheet2.getRangeByIndex(rowIndex + 1, colIndex + 1);
          cell.setText(transformedData[rowIndex][colIndex]);

          // Thêm bo viền cho ô
          cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

          if (rowIndex == 0) {
            cell.cellStyle.bold = true; // Đặt chữ in đậm
            cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
          } else {
            cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
          }
        }
      }

      final excel.Worksheet sheet3 = workbook.worksheets.add();
      sheet3.name = "Khách hàng"; // Đặt tên sheet phụ

      // Thêm dữ liệu vào sheet
      for (int rowIndex = 0; rowIndex < customerData.length; rowIndex++) {
        for (int colIndex = 0;
            colIndex < customerData[rowIndex].length;
            colIndex++) {
          final cell = sheet3.getRangeByIndex(rowIndex + 1, colIndex + 1);
          cell.setText(customerData[rowIndex][colIndex]);

          // Thêm bo viền cho ô
          cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

          if (rowIndex == 0) {
            cell.cellStyle.bold = true; // Đặt chữ in đậm
            cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
          } else {
            cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
          }
        }
      }
    }

    // Lưu file dưới dạng stream
    final List<int> bytes = workbook.saveAsStream();

    // Tải xuống file Excel trên trình duyệt
    html.AnchorElement(
        href:
            "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
      ..setAttribute("download", fileName)
      ..click();
  }

  /// Tạo và tải xuống file Excel
  static void createExcelFileCode(
      List<List<String>> dataExcel, String fileName) {
    // Tạo một workbook
    final excel.Workbook workbook = excel.Workbook();

    // Accessing worksheet via index.
    final excel.Worksheet sheet = workbook.worksheets[0];
    sheet.name = fileName; // Đổi tên sheet

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < dataExcel.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < dataExcel[rowIndex].length;
          colIndex++) {
        final cell = sheet.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(dataExcel[rowIndex][colIndex]);

        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }
    // Lưu file dưới dạng stream
    final List<int> bytes = workbook.saveAsStream();

    // Tải xuống file Excel trên trình duyệt
    html.AnchorElement(
        href:
            "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
      ..setAttribute("download", fileName)
      ..click();
  }
}
