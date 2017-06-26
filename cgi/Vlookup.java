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
	 FileInputStream myStream2 = new FileInputStream("demo3.xls");
	 NPOIFSFileSystem fs2 = new NPOIFSFileSystem(myStream2);
	 HSSFWorkbook wb2 = new HSSFWorkbook(fs2.getRoot(), true);
	 FileInputStream myStream = new FileInputStream("demo2.xls");
	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
	 
	  int xo = wbwrite.linkExternalWorkbook("demo3.xls", wb);
	 
	 HSSFSheet sheet = wb.getSheetAt(0);
	 HSSFSheet sheet2 = wb2.getSheetAt(0);
	 HSSFRow row,row2;
	 HSSFCell cell,cell2;
	 int rowStart = sheet.getFirstRowNum(),rowStart2 = sheet2.getFirstRowNum() ;
	 int rowEnd = sheet.getLastRowNum(),rowEnd2 = sheet2.getLastRowNum();
	 int fCell,lCell,fCell2,lCell2;
	 Row rowwrite[] =new Row[rowEnd+1];Row rowwrite2[] =new Row[rowEnd2+1];
	 FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();

	 Map<String,FormulaEvaluator> mapper = new HashMap<String,FormulaEvaluator>(); 
	 
	/* Collection<String> list = new ArrayList<String>();
	 list = AnalysisToolPak.getSupportedFunctionNames();
	 System.out.println(list);   */
	 
	 for(int i=rowStart;i<=rowEnd;i++){
		 row=sheet.getRow(i);
		 if(row==null){
			 System.out.println("empty accessed");
			 continue;}
		 if(row!=null){
			 rowwrite[i]=sheet_write.createRow((short)i);;
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
		    		 		rowwrite[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue());
		    		 	}
		    		 	else if(currentCell.getCellTypeEnum() == CellType.STRING){
					  		String value = currentCell.getStringCellValue();
					  		rowwrite[i].createCell(iCell).setCellValue(currentCell.getRichStringCellValue());
		    		 	}
		    	 }
		     }
				 if(i>=1){
					// try{
					 mapper.put("demo3.xls",evaluator);
					 mapper.put("demo2.xls",wb2	.getCreationHelper().createFormulaEvaluator());
					 evaluator.setupReferencedWorkbooks(mapper);
					//System.out.print(mapper);
		    		rowwrite[i].createCell(3).setCellFormula("VLOOKUP(A3,'[demo3.xls]new sheet'!$1:$65536,3,0)");
					// }catch(Exception e){}
				 }
		 }
		 FileOutputStream fileOut = new FileOutputStream("demo2.xls");
		wbwrite.write(fileOut);
		fileOut.close();
		//System.out.println("WorkBook has been created");
			  }
	 wb2.close();
	 wb.close();
	 
}

public static void main(String[] args) throws IOException{
new Vlookup();
}
}
