package Org.Data.DataComparision;
import java.io.*;
import java.lang.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class CompareExcel {
	
	public static Map < String, Object[] > compareData1 = new TreeMap < String, Object[] > ();
    public static Map < String, Object[] > compareData2 = new TreeMap < String, Object[] > ();
    public static Integer index = 1;
    public static XSSFSheet sheet1;
    public static XSSFSheet sheet2;
    
    public void fileComparision(FileInputStream excellFile1,FileInputStream excellFile2) throws IOException {
                    

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
            XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);

            // Get first/desired sheet from the workbook
            sheet1 = workbook1.getSheetAt(0);
            sheet2 = workbook2.getSheetAt(0);
//Cell mismatch
            compareData1.put(index.toString(), new Object[] {
                "ID",
                "Source Column Name",
                "Source Column Value",
                "Target Column Name",
                "Target Column Value",
                "Status"
            });
            //row missing
            compareData2.put(index.toString(), new Object[] {
                "ID",
                "Status"
            });
            // Compare sheets
            compareTwoSheets();

            // close files
            excellFile1.close();
            excellFile2.close();
            
            //Create work book and write the data
            CreateWorkBook createWorkBook = new CreateWorkBook();
            createWorkBook.writeComparisionData(compareData1, compareData2); 
                    

    }
    
    public static boolean compareHeaders(XSSFSheet sheet1,XSSFSheet sheet2) {
    	XSSFRow row1 = sheet1.getRow(0);
    	XSSFRow row2 = sheet2.getRow(0);
    	int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();
        int firstCell2= row2.getFirstCellNum();
        int lastCell2 = row2.getLastCellNum();
        boolean equalHeader=false;
        
        for (int i = firstCell1; i <= lastCell1; i++) {
            XSSFCell cell1 = row1.getCell(i);
            XSSFCell cell2 = row2.getCell(i);
            if ((cell1 != null) && (cell2 != null)) {
            	if (cell1.getStringCellValue().equals(cell2
                        .getStringCellValue())) {
                    equalHeader= true;
                } else {
                    equalHeader= false;
                    System.out.println(" Source Header Name : "+ cell1.getStringCellValue() + " Target Header Name : " + cell2.getStringCellValue());
                    
                    }
                
            }
           
        }
        
		return equalHeader;

    	
    }

    // Compare Two Sheets
    public static void compareTwoSheets() {
    	
    	if(compareHeaders(sheet1,sheet2)){
    		
    		int firstRow1 = sheet1.getFirstRowNum();
            int lastRow1 = sheet1.getLastRowNum();

            int firstRow2 = sheet2.getFirstRowNum();
            int lastRow2 = sheet2.getLastRowNum();

            
            for (int i = firstRow1 + 1; i <= lastRow1; i++) {
                XSSFRow row1 = sheet1.getRow(i);//2
                if (row1 != null) {
                    Double Row1ID = row1.getCell(0).getNumericCellValue();
                    Boolean notFound = true;
                    for (int j = firstRow2 + 1; j <= lastRow2; j++) {
                        XSSFRow row2 = sheet2.getRow(j);
                        Double Row2ID = row2.getCell(0).getNumericCellValue();
                        if (Row1ID.equals(Row2ID)) {
                            notFound =false;
                            compareTwoRows(row1, row2, i);
                            break;
                        } 
                    }
                    if(notFound){
                        index++;
                        compareData2.put(index.toString(), new Object[] {
                            Double.toString(Row1ID), "Missing row in target"
                        });
                    }
                } 
            }
            //Checking for missing rows in sheet 1
            for (int i = firstRow2 + 1; i <= lastRow2; i++) {
                XSSFRow row2 = sheet2.getRow(i);//2
                if (row2 != null) {
                    Double Row2ID = row2.getCell(0).getNumericCellValue();
                    Boolean notFound = true;
                    for (int j = firstRow1 + 1; j <= lastRow1; j++) {
                        XSSFRow row1 = sheet1.getRow(j);
                        Double Row1ID = row1.getCell(0).getNumericCellValue();
                        if (Row1ID.equals(Row2ID)) {
                            notFound =false;
                            
                            break;
                        } 
                    }
                    if(notFound){
                        index++;
                        compareData2.put(index.toString(), new Object[] {
                            Double.toString(Row2ID), "Missing row in source"
                        });
                    }
                } 
            }
    	}else {
    		System.out.println("Column headers are not equal");
    	}
        
    }

    // Compare Two Rows column
    public static void compareTwoRows(XSSFRow row1, XSSFRow row2, int rowID) {

        int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();

        // Compare all cells in a row
        for (int i = firstCell1; i <= lastCell1; i++) {
            XSSFCell cell1 = row1.getCell(i);
            XSSFCell cell2 = row2.getCell(i);
            if ((cell1 != null) && (cell2 != null)) {
                compareTwoCells(cell1, cell2, rowID, i);
            }
        }
    }

    // Compare Two Cells column
    public static void compareTwoCells(XSSFCell cell1, XSSFCell cell2, int rowID, int colID) {
        
        int type1 = cell1.getCellType();
        int type2 = cell2.getCellType();
        if (type1 == type2) {
            if (cell1.getCellStyle().equals(cell2.getCellStyle())) {
                // Compare cells based on its type
                switch (cell1.getCellType()) {
                    
                    case HSSFCell.CELL_TYPE_NUMERIC:
                        if (cell1.getNumericCellValue() != cell2
                            .getNumericCellValue()) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getNumericCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getNumericCellValue(),"Data Mismatch"
                            });
                        } 
                            
                        
                        break;
                    case HSSFCell.CELL_TYPE_STRING:
                        if (!cell1.getStringCellValue().equals(cell2
                                .getStringCellValue())) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getStringCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getStringCellValue(),"Data Mismatch"
                            });
                        } 
                        
                        break;
                    
                    case HSSFCell.CELL_TYPE_BOOLEAN:
                        if (cell1.getBooleanCellValue() != cell2
                            .getBooleanCellValue()) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getBooleanCellValue(), cell1.getBooleanCellValue(), sheet1.getRow(0).getCell(colID).getBooleanCellValue(), cell2.getBooleanCellValue(),"Data Mismatch"
                            });
                        } 
                        break;
                        
                    
                    default:
                        if (!cell1.getStringCellValue().equals(
                                cell2.getStringCellValue())) {
                        	index++;
                            compareData1.put(index.toString(), new Object[] {
                                Integer.toString(rowID), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell1.getStringCellValue(), sheet1.getRow(0).getCell(colID).getStringCellValue(), cell2.getStringCellValue(),"Data Mismatch"
                            });
                        } 
                        break;
                }
            } 
        } 
    }
}


