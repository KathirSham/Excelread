package org.test.Excel;


import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadAll {
	
	public static void main(String[] args) throws Throwable {
		
		File f = new File("C:\\Users\\DELL\\Desktop8\\Excel\\Excel1\\Account details - Copy.xlsx");
		FileInputStream stream = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("v");
		for (int i = 0; i <s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				//System.out.println(c);
				int type = c.getCellType();
				if (type==1) {
					String s1 = c.getStringCellValue();
					System.out.println(s1); 	
				}
				else if (type==0) {
					if (DateUtil.isCellDateFormatted(c)) {
						Date d = c.getDateCellValue();
					SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yy");
					String f1 = sim.format(d);
					System.out.println(f1);
					}
					else {
						double numericCellValue = c.getNumericCellValue();
						long l = (long)numericCellValue;
						String t = String.valueOf(l);
						System.out.println(t);
						
					}
					
				}
				}
			
			}
		
		
						
				
			}

	}
		
	