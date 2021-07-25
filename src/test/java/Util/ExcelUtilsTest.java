package Util;

import java.io.IOException;

public class ExcelUtilsTest
{
    public static void main(String[] args) throws IOException {
       String excelpath = "./Data/TestData.xlsx";
       String sheetname = "Sheet1";
       ExcelUtils excel = new ExcelUtils(excelpath,sheetname);

       excel.getrowcount();
       excel.getcelldata(1,0);
       excel.getcelldata(1,1);
       excel.getcelldata(1,2);
        System.out.println("!!!!!!!!!!!!!!!");
        excel.getcelldata(2,0);
        excel.getcelldata(2,1);
        excel.getcelldata(2,2);
        System.out.println("!!!!!!!!!!!!!!!");
        excel.getcelldata(3,0);
        excel.getcelldata(3,1);
        excel.getcelldata(3, 2);
    }
}
