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
	
	 Workbook wbwrite = new HSSFWorkbook();
	 Sheet sheet_write = wbwrite.createSheet("new sheet");
	 Workbook wbwrite2 = new HSSFWorkbook();
	 Sheet sheet_write2 = wbwrite.createSheet("Sheet 1");
	 
	 FileInputStream myStream = new FileInputStream("demo2.xls");
	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
	 
	 HSSFSheet sheet = wb.getSheetAt(0);
	 HSSFSheet sheet2 = wb.getSheet("Sheet 1");
	 HSSFRow row,row2;
	 HSSFCell cell,cell2;
	 int rowStart = sheet.getFirstRowNum();
	 int rowEnd = sheet.getLastRowNum();int rowEnd2 = sheet.getLastRowNum();
	 int fCell,lCell,fCell2,lCell2;
	 Row rowwrite[] =new Row[rowEnd+1];Row rowwrite2[] = new Row[rowEnd2+1];
	 
	 FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	/* Collection<String> list = new ArrayList<String>();
	 list = AnalysisToolPak.getSupportedFunctionNames();
	 System.out.println(list);   */
	 
	 for(int i=rowStart;i<=rowEnd;i++){
		 row=sheet.getRow(i);
		 if(row==null){
			 System.out.println("empty accessed");
			 continue;}
		 if(row!=null){
			 rowwrite[i]=sheet.getRow(((short)i));
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
		    		 	//	System.out.print(currentCell.getNumericCellValue() + "--");
		    		 		rowwrite[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue()+1);
		    		 	}
		    		 	else if(currentCell.getCellTypeEnum() == CellType.STRING){
					  		String value = currentCell.getStringCellValue();
					  		rowwrite[i].createCell(iCell).setCellValue(currentCell.getRichStringCellValue());
		    		 	}
		    	 }
		     }
				 if(i>=1){
					// try{
		    		//rowwrite[i].createCell(3).setCellFormula("VLOOKUP(A3,Sheet1!1:65536,3,0)");
					// }catch(Exception e){}
				 }
		 }
		 FileOutputStream fileOut = new FileOutputStream("demo2.xls");
		wbwrite.write(fileOut);
		fileOut.close();
		//System.out.println("WorkBook has been created");
			  }
	 wb.close();
	 
}

public static void main(String[] args) throws IOException{
new Vlookup();
}
}
