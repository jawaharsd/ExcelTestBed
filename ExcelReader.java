package excelProjectTestbed;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
	public static final String SAMPLE_XLSX_FILE_PATH = "my effort logger.xlsx";
    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach with lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(1);
        int counter=sheet.getPhysicalNumberOfRows();
        String A= Integer.toString(counter);
        int noOfColoumns=sheet.getRow(0).getPhysicalNumberOfCells();
        int E = noOfColoumns;


        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        Sheet sheet1 = workbook.getSheetAt(1);
        int counter1=sheet1.getPhysicalNumberOfRows();
             String B=Integer.toString(counter1);
           int noOfColoumns1=sheet1.getRow(2).getPhysicalNumberOfCells();
           int F = noOfColoumns1;

           Sheet sheet2 = workbook.getSheetAt(2);
           int counter2=sheet2.getPhysicalNumberOfRows();
                String C=Integer.toString(counter2);
              int noOfColoumns2=sheet2.getRow(0).getPhysicalNumberOfCells();
              int G = noOfColoumns2;

              Sheet sheet3 = workbook.getSheetAt(3);
              int counter3=sheet3.getPhysicalNumberOfRows();
                   String D=Integer.toString(counter3);
                 int noOfColoumns3=sheet3.getRow(0).getPhysicalNumberOfCells();
                 int H = noOfColoumns3;


        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 3. Or you can use Java 8 forEach loop with lambda
        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });
        System.out.println("Pivot table ");
        System.out.println("Number of Rows = " + A);
        System.out.println("Number of Coloumns = " + E);
        System.out.println("Effort logger ");
        System.out.println("Number of Rows = " + B);
        System.out.println("Number of Coloumns = " + F);
        System.out.println("Summary sheet ");
        System.out.println("Number of Rows = " + C);
        System.out.println("Number of Coloumns = " + G);
        System.out.println("Drop down list ");
        System.out.println("Number of Rows = " + D);
        System.out.println("Number of Coloumns = " + H);


        // Closing the workbook
        workbook.close();
    }
}
       
