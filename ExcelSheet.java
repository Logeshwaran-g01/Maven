package tests;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelSheet {

    public static void main(String[] args) {
        XSSFWorkbook workbook = null;
        //HSSFWorkbook workbook1 = null;
        try {
            FileInputStream fileInputStream = new FileInputStream("src/test/resources/Excel_1.xlsx");
            workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet =  workbook.getSheet("Sheet1");
            Map<String , Map<String , String>> testData = new HashMap<>();
            Row headerRow = sheet.getRow(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++){
                Row row = sheet.getRow(i);
                Map<String , String> colsValue = new HashMap<>();
                for (int j = 0; j < row.getLastCellNum(); j++){
                    colsValue.put(headerRow.getCell(j).getStringCellValue(), row.getCell(j).getStringCellValue());
                }
                testData.put(row.getCell(0).getStringCellValue(), colsValue);
            }
            System.out.println(testData);

        } catch (FileNotFoundException e){
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException();
        }catch (NullPointerException e){
            e.printStackTrace();
        }
    }
}
