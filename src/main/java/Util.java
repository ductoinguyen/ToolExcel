import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;

public class Util {

    public static Object getValue(Cell cell, CellType cellType) throws IOException {
        try {
            switch (cellType) {
                case DATE:
                    return cell.getDateCellValue();                                      // java.util.Date
                case HOUR:
                    return String.valueOf(cell.getDateCellValue()).split(" ")[3];   // java.lang.String
                case INTEGER:
                    return (int) cell.getNumericCellValue();                             // int
                case LONG:
                    return (long) cell.getNumericCellValue();                                // long
                case DOUBLE:
                    return cell.getNumericCellValue();                                   // double
                case STRING:
                    try {
                        return cell.getStringCellValue();                                    // java.lang.String
                    } catch (Exception e) {
                        try {
                            return String.valueOf(cell.getNumericCellValue());                                   // java.lang.String
                        } catch (Exception ex) {
                            throw e;
                        }
                    }
            }
            return "";
        } catch (Exception e) {
            System.out.println(cell.getAddress());
            FileWriter myWriter = new FileWriter("error.txt");
            myWriter.write("Lỗi ở ô: " + cell.getAddress() + ", nó cần phải là " + cellType);
            myWriter.close();
            throw e;
        }
    }

    public static String getTime(Date date, String format) {
        SimpleDateFormat formatter = new SimpleDateFormat(format);
        return formatter.format(date);
    }

    public static String getDeadline1(Date ngayCanGiao, Date ngayDatHang) {
        if (ngayCanGiao.compareTo(ngayDatHang) == 0) {
            return getTime(ngayCanGiao, "dd/MM/yyy" + " 00:00");
        }
        LocalDateTime ldt = LocalDateTime.ofInstant(ngayCanGiao.toInstant(), ZoneId.systemDefault());
        LocalDateTime date = ldt.minusDays(1);
        Date out = Date.from(date.atZone(ZoneId.systemDefault()).toInstant());
        return getTime(out, "dd/MM/yyy" + " 00:00");
    }

    public static String getDeadline2(Date ngayCanGiao, Date ngayDatHang) {
        if (ngayCanGiao.compareTo(ngayDatHang) == 0) {
            return getTime(ngayCanGiao, "dd/MM/yyy" + " 00:00");
        }
        LocalDateTime ldt = LocalDateTime.ofInstant(ngayDatHang.toInstant(), ZoneId.systemDefault());
        LocalDateTime date = ldt.plusDays(1);
        Date out = Date.from(date.atZone(ZoneId.systemDefault()).toInstant());
        return getTime(out, "dd/MM/yyy" + " 00:00");
    }

    public static String getDeadline3(Date ngayCanGiao, Date ngayCanLenMay, int soLuongMet) {
        long diff = (ngayCanGiao.getTime() - ngayCanLenMay.getTime()) / (1000 * 60 * 60 * 24);
        if (soLuongMet > 15000) {
            if (diff < 3) {
                return getTime(ngayCanGiao, "dd/MM/yyy" + " 00:00");
            } else {
                LocalDateTime ldt = LocalDateTime.ofInstant(ngayCanLenMay.toInstant(), ZoneId.systemDefault());
                LocalDateTime date = ldt.plusDays(3);
                Date out = Date.from(date.atZone(ZoneId.systemDefault()).toInstant());
                return getTime(out, "dd/MM/yyy" + " 00:00");
            }
        } else if (soLuongMet > 5000) {
            if (diff < 2) {
                return getTime(ngayCanGiao, "dd/MM/yyy" + " 00:00");
            } else {
                LocalDateTime ldt = LocalDateTime.ofInstant(ngayCanLenMay.toInstant(), ZoneId.systemDefault());
                LocalDateTime date = ldt.plusDays(2);
                Date out = Date.from(date.atZone(ZoneId.systemDefault()).toInstant());
                return getTime(out, "dd/MM/yyy" + " 00:00");
            }
        } else {
            if (diff < 1) {
                return getTime(ngayCanGiao, "dd/MM/yyy" + " 00:00");
            } else {
                LocalDateTime ldt = LocalDateTime.ofInstant(ngayCanLenMay.toInstant(), ZoneId.systemDefault());
                LocalDateTime date = ldt.plusDays(1);
                Date out = Date.from(date.atZone(ZoneId.systemDefault()).toInstant());
                return getTime(out, "dd/MM/yyy" + " 00:00");
            }
        }
    }

}
