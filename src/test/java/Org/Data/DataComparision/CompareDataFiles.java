package Org.Data.DataComparision;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.testng.annotations.Test;

public class CompareDataFiles {

	/*public static void main(String[] args) {
		
		// get input excel files
        try {
			FileInputStream excellFile1 = new FileInputStream(
			    "F:\\DataComparision\\ExcelSheet\\Source.xlsx");
			FileInputStream excellFile2 = new FileInputStream(
		            "F:\\DataComparision\\ExcelSheet\\Target.xlsx");
			CompareExcel compareExcel = new CompareExcel();
			
			compareExcel.fileComparision(excellFile1, excellFile2);
			
			System.out.println("Data Comparision Completed");
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
	}*/
	
	@Test
	public void ReadSourceFromExcel() {
		
	}
}
