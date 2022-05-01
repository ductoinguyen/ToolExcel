import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;

public class CoreBusiness {
    private Row inputRow;

    private String LSX;
    private String MA_KHACH_HANG;
    private String KINH_DOANH;
    private String MA_HANG;
    private String SO_BAN;
    private String SO_DAO;
    private String TEN_HANG;
    private int SO_LUONG_SAN_XUAT;
    private int KHO_GIAY;
    private String LOAI_GIAY;
    private Date NGAY_DAT_HANG;
    private String LOAI_MANG;
    private int SO_LUONG_MET;
    private String MAY_IN;
    private String THOI_GIAN_IN;
    private String THOI_GIAN_BE;
    private String DON_VI_TINH;
    private Date NGAY_CAN_LEN_MAY;
    private Date NGAY_CAN_GIAO;

    private boolean mustPrint;

    public CoreBusiness(Row inputRow) {
        this.inputRow = inputRow;
    }

    public void convert() throws IOException {
        for (Cell cell : inputRow) {
            int columnIndex = cell.getColumnIndex();
            int rowIndex    = cell.getRowIndex();
            int currentRow  = rowIndex + 1;

            if (currentRow < Config.ROW_BEGIN) {
                return;
            }

            if (columnIndex == 1) { // B
                NGAY_DAT_HANG = (Date) Util.getValue(cell, CellType.DATE);
            } else if (columnIndex == 2) { // C
                LSX = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 3) { // D
                MA_KHACH_HANG = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 4) { // E
                KINH_DOANH = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 5) { // F
                MA_HANG = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 6) { // G
                SO_BAN = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 7) { // H
                SO_DAO = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 8) { //
                TEN_HANG = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 9) { // J
                DON_VI_TINH = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 10) { // K
                SO_LUONG_SAN_XUAT = (Integer) Util.getValue(cell, CellType.INTEGER);
            } else if (columnIndex == 11) { // L
                LOAI_GIAY = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 12) { // M
                KHO_GIAY = (Integer) Util.getValue(cell, CellType.INTEGER);
            } else if (columnIndex == 13) { // N
                LOAI_MANG = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 14) { // O
                SO_LUONG_MET = (Integer) Util.getValue(cell, CellType.INTEGER);
            } else if (columnIndex == 15) { // P
                MAY_IN = (String) Util.getValue(cell, CellType.STRING);
            } else if (columnIndex == 16) { // Q
                THOI_GIAN_IN = (String) Util.getValue(cell, CellType.HOUR);
            } else if (columnIndex == 18) { // S
                THOI_GIAN_BE = (String) Util.getValue(cell, CellType.HOUR);
            } else if (columnIndex == 23) { // X
                NGAY_CAN_GIAO = (Date) Util.getValue(cell, CellType.DATE);
            } else if (columnIndex == 24) { // Y
                NGAY_CAN_LEN_MAY = (Date) Util.getValue(cell, CellType.DATE);
            }

        }

        mustPrint = MA_HANG.charAt(2) == 'I';

        KINH_DOANH = KINH_DOANH.trim().toUpperCase();
        HashMap<String, String> map = new HashMap<>();
        map.put("PHONG", "KDLePhong");
        map.put("LUYẾN", "KDLeHa");
        map.put("MINH", "KDMinh");
        map.put("QUANG", "KDLeQuang");
        map.put("THU", "KDNguyenThu");
        map.put("HÀ", "KDPhamHa");

        String nguoiTheoDoi = "SxQLSX, Sxkehoach, AdNguyenTuyet, KDLeHa, KDLePhong, AdNongThuy";
        if (!Arrays.asList("PHONG", "LUYẾN").contains(KINH_DOANH)) {
            nguoiTheoDoi += ", " + map.get(KINH_DOANH);
        }

        // first
        OutputSheet.addNewRow(
                LSX + " - " + MA_KHACH_HANG + " - " + MA_HANG,
                "Sxkehoach",
                "SxDongGoi",
                nguoiTheoDoi + "",
                Util.getTime(NGAY_DAT_HANG, "dd/MM/yyyy") + " 00:00",
                Util.getDeadline1(NGAY_CAN_GIAO, NGAY_DAT_HANG) + "",
                LSX + "",
                MA_KHACH_HANG + "",
                MA_HANG + "",
                SO_BAN + "",
                SO_DAO + "",
                TEN_HANG + "",
                LOAI_GIAY + "",
                LOAI_MANG + "",
                DON_VI_TINH + "",
                SO_LUONG_SAN_XUAT + "",
                KHO_GIAY + "",
                SO_LUONG_MET + "",
                MAY_IN + "",
                THOI_GIAN_IN + "",
                THOI_GIAN_BE + "");

        OutputSheet.addNewRow(
                "# Vật tư",
                "SxQLSX",
                "SxThuKhoHT",
                nguoiTheoDoi + "",
                "",
                Util.getDeadline2(NGAY_CAN_GIAO, NGAY_DAT_HANG) + "",
                LSX + "",
                MA_KHACH_HANG + "",
                MA_HANG + "",
                SO_BAN + "",
                SO_DAO + "",
                TEN_HANG + "",
                LOAI_GIAY + "",
                LOAI_MANG + "",
                DON_VI_TINH + "",
                SO_LUONG_SAN_XUAT + "",
                KHO_GIAY + "",
                SO_LUONG_MET + "",
                MAY_IN + "",
                THOI_GIAN_IN + "",
                THOI_GIAN_BE + "");

        // in
        if (mustPrint) {
            OutputSheet.addNewRow(
                    "# In",
                    "SxQLSX",
                    "SxInHT",
                    nguoiTheoDoi + "",
                    "",
                    Util.getDeadline3(NGAY_CAN_GIAO, NGAY_CAN_LEN_MAY, SO_LUONG_MET) + "",
                    LSX + "",
                    MA_KHACH_HANG + "",
                    MA_HANG + "",
                    SO_BAN + "",
                    SO_DAO + "",
                    TEN_HANG + "",
                    LOAI_GIAY + "",
                    LOAI_MANG + "",
                    DON_VI_TINH + "",
                    SO_LUONG_SAN_XUAT + "",
                    KHO_GIAY + "",
                    SO_LUONG_MET + "",
                    MAY_IN + "",
                    THOI_GIAN_IN + "",
                    THOI_GIAN_BE + "");
        }

        OutputSheet.addNewRow(
                "# Bế Xẻ",
                "SxQLSX",
                "SxBeXeHT",
                nguoiTheoDoi + "",
                "",
                Util.getDeadline3(NGAY_CAN_GIAO, NGAY_CAN_LEN_MAY, SO_LUONG_MET) + "",
                LSX + "",
                MA_KHACH_HANG + "",
                MA_HANG + "",
                SO_BAN + "",
                SO_DAO + "",
                TEN_HANG + "",
                LOAI_GIAY + "",
                LOAI_MANG + "",
                DON_VI_TINH + "",
                SO_LUONG_SAN_XUAT + "",
                KHO_GIAY + "",
                SO_LUONG_MET + "",
                MAY_IN + "",
                THOI_GIAN_IN + "",
                THOI_GIAN_BE + "");

        OutputSheet.addNewRow(
                "# QC",
                "SxQLSX",
                "SxQCHT",
                nguoiTheoDoi + "",
                "",
                Util.getDeadline1(NGAY_CAN_GIAO, NGAY_DAT_HANG) + "",
                LSX + "",
                MA_KHACH_HANG + "",
                MA_HANG + "",
                SO_BAN + "",
                SO_DAO + "",
                TEN_HANG + "",
                LOAI_GIAY + "",
                LOAI_MANG + "",
                DON_VI_TINH + "",
                SO_LUONG_SAN_XUAT + "",
                KHO_GIAY + "",
                SO_LUONG_MET + "",
                MAY_IN + "",
                THOI_GIAN_IN + "",
                THOI_GIAN_BE + "");

        OutputSheet.addNewRow(
                "# Đóng gói",
                "SxQLSX",
                "SxDongGoi",
                nguoiTheoDoi + "",
                "",
                Util.getDeadline1(NGAY_CAN_GIAO, NGAY_DAT_HANG) + "",
                LSX + "",
                MA_KHACH_HANG + "",
                MA_HANG + "",
                SO_BAN + "",
                SO_DAO + "",
                TEN_HANG + "",
                LOAI_GIAY + "",
                LOAI_MANG + "",
                DON_VI_TINH + "",
                SO_LUONG_SAN_XUAT + "",
                KHO_GIAY + "",
                SO_LUONG_MET + "",
                MAY_IN + "",
                THOI_GIAN_IN + "",
                THOI_GIAN_BE + "");
    }
}
