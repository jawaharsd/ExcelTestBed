package excelWriterExample;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class MyEffortLogger {
     public static void main1(String[] args) {
            FileOutputStream copy = null;
            try {
                copy = new FileOutputStream(new File("Copy of EL.xlsx"));
                
                String SAMPLE_XLSX_FILE_PATH = "My effort logger.xlsx";
                Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
                workbook.write(copy);
                workbook.close();
                copy.flush();
            }
            catch (Exception e) {
                e.printStackTrace();
            }
     }
}
