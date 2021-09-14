package frameWork;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Sample {

	public static void main(String[] args) throws IOException {
		
		try {
			File excelLoc = new File("C:\\Users\\dell\\eclipse-workspace\\MyProject\\TestData\\Book1.xlsx");
			
			FileInputStream stream = new FileInputStream(excelLoc);
			
			XSSFWorkbook w = new XSSFWorkbook(stream);
		  XSSFSheet s = w.getSheet("Sheet1");
	      
	
	
		}
		}
		
	   for (int i = 0; i <s.getPhysicalNumberOfRows(); i++) {
		XSSFRow r = s.getRow(i);
		
		for (int j = 0; j <r.getPhysicalNumberOfCells(); j++) {
			XSSFCell c = r.getCell(j);
			System.out.println(c);
			
		}
		
		
	}
	  
		
		
	}

}
