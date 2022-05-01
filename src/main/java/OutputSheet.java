import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class OutputSheet {

    public static List<List<String>> schema = Arrays.asList(
        Arrays.asList("A",  "0",  "Tên công việc",      "12330"),
        Arrays.asList("B",  "1",  "Người giao",         "2520"),
        Arrays.asList("C",  "2",  "Người thực hiện",    "3700"),
        Arrays.asList("D",  "3",  "Người theo dõi",     "13100"),
        Arrays.asList("E",  "4",  "Khẩn cấp",           "4400"),
        Arrays.asList("F",  "5",  "Quan trọng",         "4400"),
        Arrays.asList("G",  "6",  "Danh sách nhãn",     "4400"),
        Arrays.asList("H",  "7",  "Ngày bắt đầu",       "4400"),
        Arrays.asList("I",  "8",  "Thời hạn",           "4400"),
        Arrays.asList("J",  "9",  "Hoàn thành thực tế", "4400"),
        Arrays.asList("K",  "10", "Mô tả công việc",    "4400"),
        Arrays.asList("L",  "11", "Trạng thái",         "4400"),
        Arrays.asList("M",  "12", "Kết quả công việc",  "4400"),
        Arrays.asList("N",  "13", "Mục tiêu",           "4400"),
        Arrays.asList("O",  "14", "Lệnh SX",            "4400"),
        Arrays.asList("P",  "15", "Mã KH",              "4400"),
        Arrays.asList("Q",  "16", "Mã hàng",            "6630"),

        Arrays.asList("R",  "17", "Số bản",             "4400"),
        Arrays.asList("S",  "18", "Số dao",             "4400"),

        Arrays.asList("T",  "19", "Tên hàng",           "16700"),
        Arrays.asList("U",  "20", "Mã giấy vật tư",     "4400"),
        Arrays.asList("V",  "21", "Màng",               "4400"),
        Arrays.asList("W",  "22", "Đơn VT",             "4400"),
        Arrays.asList("X",  "23", "Số lượng cần SX",    "4400"),
        Arrays.asList("Y",  "24", "Khổ giấy",           "4400"),
        Arrays.asList("Z",  "25", "slg mét dài",        "4400"),
        Arrays.asList("AA", "26", "Máy In",             "4400"),
        Arrays.asList("AB", "27", "TG In",              "4400"),
        Arrays.asList("AC", "28", "TG Bế",              "4400"),
        Arrays.asList("AD", "29", "Ngày tạo",           "4400"),
        Arrays.asList("AE", "30", "Mã công việc (ID)",  "4400")
    );

    public static Sheet sheet;
    public static Workbook workbook;
    public static int currentRow = 0;

    public static void create() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet(Config.OUTPUT_SHEET_NAME);
        for (List<String> col : schema) {
            sheet.setColumnWidth(Integer.parseInt(col.get(1)), Integer.parseInt(col.get(3)));
        }

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        font.setBold(false);
        headerStyle.setFont(font);

        for (List<String> col : schema) {
            Cell headerCell = header.createCell(Integer.parseInt(col.get(1)));
            headerCell.setCellValue(col.get(2));
            headerCell.setCellStyle(headerStyle);
        }
    }

    public static void addNewRow(
            String tenCongViec,         // A    0   x
            String nguoiGiao,           // B    1   x
            String nguoiThucHien,       // C    2   x
            String nguoiTheoDoi,        // D    3   x
            String ngayBatDau,          // H    7   x
            String thoiHan,             // I    8   x
            String lenhSanXuat,         // O    14  x
            String maKhachHang,         // P    15  x
            String maHang,              // Q    16  x
            String soBan,               // R    17
            String soDao,               // S    18
            String tenHang,             // T    19  x
            String maGiayVatTu,         // U    20  x
            String mang,                // V    21  x
            String donViTinh,           // W    22
            String soLuongSanXuat,      // X    23
            String khoGiay,             // Y    24
            String soLuongMet,          // Z    25
            String mayIn,               // AA   26
            String thoiGianIn,          // AB   27
            String thoiGianBe           // AC   28
    ) {
        Row row = sheet.createRow(++currentRow);

        Cell cell = row.createCell(0);
        cell.setCellValue(tenCongViec);

        cell = row.createCell(1);
        cell.setCellValue(nguoiGiao);

        cell = row.createCell(2);
        cell.setCellValue(nguoiThucHien);

        cell = row.createCell(3);
        cell.setCellValue(nguoiTheoDoi);

        row.createCell(4);
        row.createCell(5);
        row.createCell(6);

        cell = row.createCell(7);
        cell.setCellValue(ngayBatDau);

        cell = row.createCell(8);
        cell.setCellValue(thoiHan);

        row.createCell(9);
        row.createCell(10);
        row.createCell(11);
        row.createCell(12);
        row.createCell(13);

        cell = row.createCell(14);
        cell.setCellValue(lenhSanXuat);

        cell = row.createCell(15);
        cell.setCellValue(maKhachHang);

        cell = row.createCell(16);
        cell.setCellValue(maHang);

        cell = row.createCell(17);
        cell.setCellValue(soBan);

        cell = row.createCell(18);
        cell.setCellValue(soDao);

        cell = row.createCell(19);
        cell.setCellValue(tenHang);

        cell = row.createCell(20);
        cell.setCellValue(maGiayVatTu);

        cell = row.createCell(21);
        cell.setCellValue(mang);

        cell = row.createCell(22);
        cell.setCellValue(donViTinh);

        cell = row.createCell(23);
        cell.setCellValue(soLuongSanXuat);

        cell = row.createCell(24);
        cell.setCellValue(khoGiay);

        cell = row.createCell(25);
        cell.setCellValue(soLuongMet);

        cell = row.createCell(26);
        cell.setCellValue(mayIn);

        cell = row.createCell(27);
        cell.setCellValue(thoiGianIn);

        cell = row.createCell(28);
        cell.setCellValue(thoiGianBe);

        row.createCell(29);
        row.createCell(30);
    }

    public static void end() throws IOException {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + Config.OUTPUT_FILE_NAME;

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();
    }
}
