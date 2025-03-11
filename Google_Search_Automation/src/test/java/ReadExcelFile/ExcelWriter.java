package ReadExcelFile;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {
    private static final String FILE_PATH = "search_results.xlsx";

    public static void writeToExcel(String keyword, String longest, String shortest) {
        try {
            Workbook workbook;
            Sheet sheet;

            File file = new File("C:\\Usersbasicjavajavabasicm1\\Google_Search_Automation\\Excel\\Test.xls");

            // এক্সেল ফাইল যদি আগে থাকে, তাহলে ওপেন করো, না থাকলে নতুন বানাও
            if (file.exists()) {
                FileInputStream fis = new FileInputStream(file);
                workbook = new HSSFWorkbook(fis);
                sheet = workbook.getSheetAt(0);
                fis.close();
            } else {
                workbook = new HSSFWorkbook();
                sheet = workbook.createSheet("Search Results");

                // হেডার লেখো (প্রথমবার এক্সেল ফাইল বানালে)
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Keyword");
                headerRow.createCell(1).setCellValue("Longest Suggestion");
                headerRow.createCell(2).setCellValue("Shortest Suggestion");
            }

            // নতুন ডাটা অ্যাড করো
            int rowCount = sheet.getLastRowNum();
            Row row = sheet.createRow(rowCount + 1);
            row.createCell(0).setCellValue(keyword);
            row.createCell(1).setCellValue(longest);
            row.createCell(2).setCellValue(shortest);

            // এক্সেল ফাইল সংরক্ষণ করো
            FileOutputStream fos = new FileOutputStream(FILE_PATH);
            workbook.write(fos);
            fos.close();
            workbook.close();

            System.out.println("Data written to Excel successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
