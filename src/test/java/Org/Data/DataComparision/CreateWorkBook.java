package Org.Data.DataComparision;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkBook {

	
	 public  void writeComparisionData(Map < String, Object[] > compareData1,Map < String, Object[] > compareData2) throws IOException {
	    	//Create blank workbook
	        XSSFWorkbook workbook = new XSSFWorkbook();

	        //Create a blank sheet
	        XSSFSheet spreadsheet1 = workbook.createSheet(" Data Mismatch ");
	        XSSFSheet spreadsheet2 = workbook.createSheet(" Data Missing ");

	        //Create row object
	        XSSFRow row;

	        //Iterate over data and write to sheet
	        Set < String > keyid1 = compareData1.keySet();
	        int rowid = 0;

	        for (String key: keyid1) {
	            row = spreadsheet1.createRow(rowid++);
	            Object[] objectArr = compareData1.get(key);
	            int cellid = 0;

	            for (Object obj: objectArr) {
	                Cell cell = row.createCell(cellid++);
	                cell.setCellValue(String.valueOf(obj));
	            }
	        }

	        //Iterate over data and write to sheet
	        Set < String > keyid2 = compareData2.keySet();
	        rowid = 0;

	        for (String key: keyid2) {
	            row = spreadsheet2.createRow(rowid++);
	            Object[] objectArr = compareData2.get(key);
	            int cellid = 0;

	            for (Object obj: objectArr) {
	                Cell cell = row.createCell(cellid++);
	                cell.setCellValue(String.valueOf(obj));
	            }
	        }
	        //Write the workbook in file system
	        FileOutputStream out = new FileOutputStream(new File("F:\\DataComparision\\ExcelSheet\\Writesheet.xlsx"));

	        workbook.write(out);


	    	
	    }
}
