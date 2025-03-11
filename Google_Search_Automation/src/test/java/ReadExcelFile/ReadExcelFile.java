package ReadExcelFile;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReadExcelFile {
    public static void main (String[] args) throws IOException {
        File file = new File("C:\\Usersbasicjavajavabasicm1\\Google_Search_Automation\\Excel\\Test.xls");
        FileInputStream fis = new FileInputStream(file);

        HSSFWorkbook workbook = new HSSFWorkbook(fis); //old version excel er jonno HSSF r new er jonno XSSF
        HSSFSheet sheet = workbook.getSheetAt(1);
     //  String cellValue = sheet.getRow(0).getCell(0).getStringCellValue();
      // System.out.println(cellValue);

        int rowCount = sheet.getPhysicalNumberOfRows();
        for ( int i=0; i< rowCount; i++){

            HSSFRow row = sheet.getRow(i);

            int cellCount = row.getPhysicalNumberOfCells();
            for ( int j=0; j< cellCount; j++){
                HSSFCell cell = row.getCell(j);
                String cellValue = getCellValue(cell);
                System.out.print("||"+cellValue);
            }
            System.out.println();
        }


        workbook.close();
        fis.close();

    }
public static String getCellValue(HSSFCell cell){
        switch (cell.getCellType()){
            case NUMERIC :
                return String.valueOf(cell.getNumericCellValue());

            case BOOLEAN :
            return String.valueOf(cell.getBooleanCellValue());
            case STRING :
            return cell.getStringCellValue();
            default:cell.getStringCellValue();
        }
    return null;
}
}
