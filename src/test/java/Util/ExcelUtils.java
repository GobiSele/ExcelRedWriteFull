package Util;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;

public class ExcelUtils {

    static XSSFWorkbook workbook;
    static  XSSFSheet sheet;

    public ExcelUtils(String excelpath, String sheetname) {
        try {
            workbook = new XSSFWorkbook(excelpath);
            sheet = workbook.getSheet(sheetname);
            }
        catch (Exception exp) {
            System.out.println(exp.getCause());
            System.out.println(exp.getMessage());
            exp.getStackTrace();
                            }
    }

    public static void getcelldata (int rownum, int colunum) throws IOException {
            DataFormatter formatter = new DataFormatter();
            Object value = formatter.formatCellValue(sheet.getRow(rownum).getCell(colunum));
            //double value = sheet.getRow(1).getCell(2).getNumericCellValue();
            System.out.println(value);
                                                         }
        public static void getrowcount ()
        {
                int rowcount = sheet.getPhysicalNumberOfRows();
                System.out.println("No of Rows " + rowcount);

         }
}
