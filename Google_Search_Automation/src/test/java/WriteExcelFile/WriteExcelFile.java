package WriteExcelFile;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelFile {

    public static void main (String[] args) throws IOException {

      HSSFWorkbook Workbook = new HSSFWorkbook();
        HSSFSheet sheet = Workbook.createSheet();
        sheet.createRow(0);
        sheet.getRow(0).createCell(1).setCellValue("hello");
        sheet.getRow(0).createCell(0).setCellValue("world");

        sheet.createRow(1);
        sheet.getRow(1).createCell(0).setCellValue("hi");
        sheet.getRow(1).createCell(1).setCellValue("dana");

        File file = new File("C:\\Usersbasicjavajavabasicm1\\Google_Search_Automation\\Excel\\Test.xls");

       // workbook.write(file);
        FileOutputStream fileOut = new FileOutputStream(file);
        Workbook.write(fileOut);
        fileOut.close();  // ফাইল বন্ধ করা জরুরি


    }
}
