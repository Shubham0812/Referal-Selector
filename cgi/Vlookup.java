package cgi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Vlookup {
	
Vlookup() throws IOException{
	
	 Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
	    Sheet sheet1 = wb.createSheet("first sheet");
	    Sheet sheet2 = wb.createSheet("second sheet");
	    
	    FileInputStream myStream = new FileInputStream("demo2.xls");
	    NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	    HSSFWorkbook wbread = new HSSFWorkbook(fs.getRoot(), true);
	    HSSFSheet sheetx  = wbread.getSheetAt(0);
	    HSSFSheet sheetx2 = wbread.getSheet("second sheet");
	    HSSFRow row,row2;
	    HSSFCell cell,cell2;
   int counter = 1;
   int rowStart = sheetx.getFirstRowNum() ;int rowStart2 = sheetx2.getFirstRowNum();
   int rowEnd = sheetx.getLastRowNum() ;int rowEnd2 = sheetx2.getLastRowNum();
   int fCell,fCell2,lCell,lCell2;
   Row rowwrite[] =new Row[rowEnd+1];Row rowwrite2[] =new Row[rowEnd2+1];
   for(int i=rowStart;i<=rowEnd;i++){
		 row=sheetx.getRow(i);
		 if(row==null){
	    		System.out.println("empty accessed");
			continue;
			}
	 if(row!=null){
		 rowwrite[i]=sheet1.createRow((short)i);;
		 System.out.println("Last cell : " +row.getLastCellNum());
			 fCell = row.getFirstCellNum(); 
	     lCell = row.getLastCellNum();
	     
	     for (int iCell = fCell; iCell < lCell; iCell++) {
	    	 cell = row.getCell(iCell);
	    	 if(cell==null){
	    		 continue; 
	    	 }
	    	 else{
	    		 Cell currentCell = cell;
	    		 	 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
	    		 		System.out.print(currentCell.getNumericCellValue() + "--");
	    		 		rowwrite[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue());
	    		 	}
	    		 	else if(currentCell.getCellTypeEnum() == CellType.STRING){
				  		String value = currentCell.getStringCellValue();
				  		rowwrite[i].createCell(iCell).setCellValue(currentCell.getRichStringCellValue());
	    		 	}
	    	 }
	     }
	 }
	 rowwrite[0].createCell(3).setCellValue("Missing values");
	 rowwrite[i].createCell(3).setCellFormula("VLOOKUP(A"+counter+",'second sheet'!1:65536,3,0)");
	 counter++;
   }
   
   for(int i=rowStart2;i<=rowEnd2;i++){
		 row2=sheetx2.getRow(i);
		 if(row2==null){
	    		System.out.println("empty accessed");
			continue;
			}
	 if(row2!=null){
		 rowwrite2[i]=sheet2.createRow((short)i);;
		 System.out.println("Last cell : " +row2.getLastCellNum());
			 fCell2 = row2.getFirstCellNum(); 
	     lCell2 = row2.getLastCellNum();
	     
	     for (int iCell = fCell2; iCell < lCell2; iCell++) {
	    	 cell = row2.getCell(iCell);
	    	 if(cell==null){
	    		 continue; 
	    	 }
	    	 else{
	    		 Cell currentCell = cell;
	    		 	 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
	    		 		System.out.print(currentCell.getNumericCellValue() + "--");
	    		 		rowwrite2[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue());
	    		 	}
	    		 	else if(currentCell.getCellTypeEnum() == CellType.STRING){
				  		String value = currentCell.getStringCellValue();
				  		rowwrite2[i].createCell(iCell).setCellValue(currentCell.getRichStringCellValue());
	    		 	}
	    	 }
	     }
	 }
}
   	//sheet2.createRow(0).createCell(0).setCellValue("ID");
	FileOutputStream fileOut = new FileOutputStream("demo2.xls");
	 wb.write(fileOut);;
	fileOut.close();
	System.out.println("WorkBook has been created");
   }
	
	                        

public static void main(String[] args) throws IOException{
new Vlookup();
}
}
