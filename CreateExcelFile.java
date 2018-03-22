package excelWriterExample;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.BorderStyle;
/*
 * Here we will learn how to create Excel file and header for the same.
 */
public class CreateExcelFile {
	
	int rownum = 0;
	XSSFSheet firstSheet;
	Collection<File> files;
	XSSFWorkbook workbook;
	File exactFile;
	{
		workbook = new XSSFWorkbook();
	    XSSFFont font = workbook.createFont();
	    font.setFontHeightInPoints((short) 50);
	    font.setFontName("Arial");
	    font.setItalic(true);

		firstSheet = workbook.createSheet("First Sheet");
		Row headerRow = firstSheet.createRow(rownum);
		headerRow.setHeightInPoints(40);
	}
	public static void main(String args[]) throws Exception {
		List<String> headerRow = new ArrayList<String>();
		headerRow.add("Employee No");
		headerRow.add("Employee Name");
		headerRow.add("Employee Address");
		List<String> firstRow = new ArrayList<String>();
		firstRow.add("1111");
		firstRow.add("Gautam");
		firstRow.add("India");
		List<String> secondRow = new ArrayList<String>();
		secondRow.add("2222");
		secondRow.add("Lynn");
		secondRow.add("USA");
		List<String> thirdRow = new ArrayList<String>();
		thirdRow.add("3333");
		thirdRow.add("Amer");
		thirdRow.add("Jordan");
		List<List<String>> rows = new ArrayList<List<String>>();
		rows.add(headerRow);
		rows.add(firstRow);
		rows.add(secondRow);
		rows.add(thirdRow);
		CreateExcelFile cls = new CreateExcelFile(rows);
		cls.createExcelFile();
	}
	void createExcelFile(){
		FileOutputStream fos = null;
		try {
			fos=new FileOutputStream(new File("ExcelSheet.xlsx"));
			XSSFCellStyle hsfstyle=workbook.createCellStyle();
			hsfstyle.setBorderBottom(BorderStyle.THICK);
			XSSFColor myColor = new XSSFColor(Color.BLUE);
			hsfstyle.setFillForegroundColor(myColor);
			hsfstyle.setFillBackgroundColor(myColor);
			hsfstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			workbook.write(fos);
			workbook.close();
			fos.flush();
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	CreateExcelFile(List<List<String>> rowList) throws Exception {
		try {
			for (int j = 0; j < rowList.size(); j++) {
				Row row = firstSheet.createRow(rownum);
				List<String> colList= rowList.get(j);
				for(int k=0; k<colList.size(); k++)
				{
					Cell cell = row.createCell(k);
					cell.setCellValue(colList.get(k));
				}
				rownum++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
	}
}