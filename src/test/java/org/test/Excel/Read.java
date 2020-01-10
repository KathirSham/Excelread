package org.test.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {

	public static void main(String[] args) throws Exception {
		File location = new File("C:\\Users\\DELL\\Desktop8\\Excel\\Excel\\Account details - Copy.xlsx");
		
		//FileInputStream stream=new FileInputStream(location);
		
		FileInputStream stream = new FileInputStream(location);
		
		Workbook wb= new XSSFWorkbook(stream);
		
		Sheet s = wb.getSheet("sheet1");
		
		
		//Row r= s.getRow(1);
		
		//Cell c =r.getCell();
				
		//int type =c.getCellType();
			
		//System.out.println(c);
		
		for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
			
			Row r = s.getRow(i);
		
		
		for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
			
			Cell c = r.getCell(j);
			
			System.out.println(c);
		}
		
		
		
	}
	}
}

