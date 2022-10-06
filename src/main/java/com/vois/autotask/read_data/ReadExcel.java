package com.vois.autotask.read_data;
import org.apache.commons.mail.EmailException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sikuli.script.Screen;
import java.io.FileInputStream;
import java.io.IOException;



public class ReadExcel {
    static String excel = "src/main/resources/Data.xlsx";
    private static final Screen screen = new Screen();
    public static void main(String[] args) throws EmailException {

        try {
            // Open the Excel file
            FileInputStream myFile = new FileInputStream(excel);
            // Access the required test data sheet
            XSSFWorkbook xwb = new XSSFWorkbook(myFile);
            XSSFSheet sheet = xwb.getSheet("sheet1");
            int row_count = sheet.getPhysicalNumberOfRows();
            int column_count = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 0; i < row_count; i++) {
                for (int j = 1; j < column_count; j++) {
//                    System.out.println(sheet.getRow(i).getCell(j).getStringCellValue()+"|");
                    System.out.println(sheet.getRow(i).getCell(j));
                }
                System.out.println("\n");
            }
        } catch (IOException e) {
            System.out.println("Test data file not found");
        }
    }
}




//        System.out.println(FOLDER);
//        try{
//            String path = "img/sort.png";
//            Pattern target = new Pattern(path);
//            Screen scr = new Screen();
//            scr.click(target);
//        }
//        catch(Exception e)
//        {
//            e.printStackTrace();
//        }
