import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {

    public static void lock() {

    }

    public static void main(String[] args) throws IOException {
        try {
            Preparation.run();
            FileInputStream file = new FileInputStream(Config.INPUT_FILE_NAME);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(Config.INPUT_SHEET_NAME);
            OutputSheet.create();
            int i = 1;
            for (Row row : sheet) {
                if (i < Config.ROW_BEGIN) {
                    i++;
                    continue;
                }
                if (i > Config.ROW_END) {
                    break;
                }
                CoreBusiness coreBusiness = new CoreBusiness(row);
                coreBusiness.convert();
                i++;
            }
            OutputSheet.end();
        } catch (Exception e) {
            FileWriter myWriter = new FileWriter("error.txt");
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            String sStackTrace = sw.toString(); // stack trace as a string
            System.out.println(sStackTrace);
            myWriter.write(sStackTrace);
            myWriter.close();
        }
    }
}
