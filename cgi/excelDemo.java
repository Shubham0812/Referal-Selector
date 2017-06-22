package cgi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class excelDemo {
public static void main(String[] args) throws IOException{
	Workbook wb = new HSSFWorkbook();
	 CreationHelper createHelper = wb.getCreationHelper();
	 Sheet sheet = wb.createSheet("new sheet");
	 Row row = sheet.createRow((short)0);
	 Row row2 = sheet.createRow((short)2);
	 Cell cell = row.createCell(0);
	 System.out.println(cell.getAddress());
	       cell.setCellValue(1);
	       row.createCell(1).setCellValue(1.2);
	       row.createCell(2).setCellValue("Hello I have been instantiated with Java Code");
	       sheet.setColumnWidth(10,1100*4);
	       
	      //cell style to set date format
	     /*  CellStyle style = wb.createCellStyle();
	       Font font = wb.createFont();
	       font.setFontHeightInPoints((short)11);
	       font.setFontName(HSSFFont.FONT_ARIAL);
	       font.setBold(true);
	       style.setFont(font); */
	       CellStyle style = wb.createCellStyle();
	       style.setDataFormat(
	           createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
	       
	       Cell cell2 = row2.createCell(3);
	            cell2.setCellValue(new Date());
	            cell2.setCellStyle(style);
	       Cell cell3 = row2.createCell(5);
    	   		cell3.setCellValue(2332721);
	       Cell cell4 = row2.createCell(7);
	    		cell4.setCellValue("J0517-0226");	
	       Cell cell5 = row2.createCell(7);
		        cell5.setCellValue("917760838432");	
		        cell5.setCellStyle(style);
	      //creating a new cell and applying concatenation
	       
		        Cell celler = row2.createCell(10);
	            celler.setCellFormula("CONCATENATE("+cell3.getAddress()+","+cell4.getAddress()+")");
	      
	       //writing the wb
	       FileOutputStream fileOut = new FileOutputStream("demo.xls");
	       wb.write(fileOut);
	       fileOut.close();
	       System.out.println("wb has been created");
	       
}
}
