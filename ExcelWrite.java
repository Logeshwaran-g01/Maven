import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWrite {
    public static void main(String[] args) throws IOException {
        //Create a Workbook
        Workbook workbook = new XSSFWorkbook();
        // Create a Sheet
        Sheet sheet = workbook.createSheet("Sheet1");
        //Create a Row and put some cells in it
        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("Name");
        Cell headerCel2 = headerRow.createCell(1);
        headerCel2.setCellValue("Age");
        Cell heaederCell3 = headerRow.createCell(2);
        heaederCell3.setCellValue("Email");

        //Data
        Object[][] obj ={
                {"John Doe", 28, "john@test.com"},
                {"Jane Doe", 30, "johnn@test.com"},
                {"Bob Smith", 37, "Jacky@example.com"},
                {"Swapnil", 35, "swapnil@example.com"}
        };

        int rowNum = 1;
        for (Object[] rowData:obj ){
            Row row = sheet.createRow(rowNum++);

            int colNum = 0;
            for (Object field : rowData){
                Cell cell = row.createCell(colNum++);
                if (field instanceof  String){
                    cell.setCellValue((String)field);
                }else if (field instanceof Integer){
                    cell.setCellValue((Integer)field);
                }
            }
        }

        //Write the output file
        try {
            FileOutputStream fos = new FileOutputStream("src/test/resources/Excel_2.xlsx");
            workbook.write(fos);
            System.out.println("Excel File created Successfully");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        finally {
            workbook.close();
        }
    }
}
