package com.CompareTwoExcelSheets.org;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

public class CompareResponseTimeInExcelSheet    
{
	public static void main(String[] args) throws IOException {

	FileInputStream file1 = new FileInputStream("./src/test/resources/Data1.xlsx");
	FileInputStream file2 = new FileInputStream("./src/test/resources/Data2.xlsx");
	@SuppressWarnings("resource")
	XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
	@SuppressWarnings("resource")
	XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
	XSSFSheet sheet1 = workbook1.getSheetAt(0);
	XSSFSheet sheet2 = workbook2.getSheetAt(0);
	compareDataInSheet(sheet1,sheet2);
	file1.close();
	file2.close();
	}
	public static void compareDataInSheet(XSSFSheet sheet1 ,XSSFSheet sheet2) {
		for (int j = 0; j <= sheet1.getLastRowNum(); j++) {
            if (sheet2.getLastRowNum() <= j) {
                return;
            }

            XSSFRow row1 = sheet1.getRow(j);
            XSSFRow row2 = sheet2.getRow(j);

            if ((row1 == null) || (row2 == null)) {
                continue;
            }

            compareDataInRow(row1,row2);
        }
	}
	public static void compareDataInRow(XSSFRow row1,XSSFRow row2) {
		//System.out.println(row1.getLastCellNum());
        //for (int k = 0; k <= row1.getLastCellNum(); k++) {
        	int k=0 ;
        	while(k<=2) {
            if (row2.getLastCellNum() <= k) {
                return;
            }

            XSSFCell cell1 = row1.getCell(k);
            XSSFCell cell2 = row2.getCell(k);

            if ((cell1 == null) || (cell2 == null)) {
                continue;
            }

            compareDataInCell(cell1, cell2 , k);
            k++ ; 
        }
    }
	public static void compareDataInCell(XSSFCell cell1,XSSFCell cell2 , int k) {
		switch(cell1.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			isCellContentMatchesForNumeric(cell1,cell2);
			break;
		case Cell.CELL_TYPE_STRING:
			isCellContentMatchesForString(cell1,cell2);
			break;
		}		
	}
	private static void isCellContentMatchesForNumeric(XSSFCell cell1, XSSFCell cell2) {
		// TODO Auto-generated method stub
		double num1 = cell1.getNumericCellValue();
		double num2 = cell2.getNumericCellValue();
				if(num1 < num2) {
					System.out.println("Low Response Time in Previous Build");
					System.out.println(num1);
				}
					else {
					System.out.println("Low Response Time in Current Build");
					System.out.println(num2);
				}
		}
	private static void isCellContentMatchesForString(XSSFCell cell1, XSSFCell cell2) {
		// TODO Auto-generated method stub
		String str1 = cell1.getStringCellValue();
		String str2 = cell2.getStringCellValue();
		System.out.println(str1);
		System.out.println(str2);
	}    
}

//public static void compareDataInCell(XSSFCell cell1,XSSFCell cell2) {
//switch (cell1.getCellType()) {
//case Cell.CELL_TYPE_NUMERIC:
//	System.out.println(cell1.getNumericCellValue());
//	break;
//case Cell.CELL_TYPE_STRING:
//	System.out.println(cell1.getStringCellValue());
//	break;
//}
//
//switch (cell2.getCellType()) {
//case Cell.CELL_TYPE_NUMERIC:
//	System.out.println(cell2.getNumericCellValue());
//	break;
//case Cell.CELL_TYPE_STRING:
//	System.out.println(cell2.getStringCellValue());
//	break;
//}
//
//}
