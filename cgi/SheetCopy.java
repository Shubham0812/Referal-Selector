package cgi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SheetCopy {
SheetCopy() throws IOException{
	 Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
	 Sheet sheet1 = wb.createSheet("Sheet1");
	 Sheet sheet2 = wb.createSheet("Sheet2");
	 
	 FileInputStream myStream = new FileInputStream("Njoyn Master Tracker-16-18.xls(formatted).xls");
	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	 HSSFWorkbook wbread = new HSSFWorkbook(fs.getRoot(), true);
	 
	 FileInputStream myStream2 = new FileInputStream("Candidate Referrals (Generic).xls");
	 NPOIFSFileSystem fs2 = new NPOIFSFileSystem(myStream2);
	 HSSFWorkbook wbread2 = new HSSFWorkbook(fs2.getRoot(), true);	
	 CreationHelper createHelper = wb.getCreationHelper();
	 HSSFSheet sheetx  = wbread.getSheetAt(0);
	 HSSFSheet sheetx2 = wbread2.getSheetAt(0);
	 HSSFRow row,row2;
	    int counter1 = 7;
	    int counter2 = 7;
	 HSSFCell cell,cell2;
	
	   int rowStart = sheetx.getFirstRowNum() ;int rowStart2 = sheetx2.getFirstRowNum();
	   int rowEnd = sheetx.getLastRowNum() ;int rowEnd2 = sheetx2.getLastRowNum();
	   int fCell,fCell2,lCell,lCell2;
	   FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
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
		    		 	else if(currentCell.getCellTypeEnum()==CellType.FORMULA){
		    		 		rowwrite[i].createCell(iCell).setCellValue(12);
		    		 	}
		    	 }
		     }
		 }
		 //rowwrite[0].createCell(3).setCellValue("Missing values");
		 //rowwrite[i].createCell(3).setCellFormula("VLOOKUP(A"+counter+",'second sheet'!1:65536,3,0)");
		 //counter++;
	   }
	   
		for(int i=rowStart2;i<=rowEnd2;i++){
			row2=sheetx2.getRow(i);
			if(row2==null){
				System.out.println("empty accessed");
				continue;
				}
			if(row2!=null){
				rowwrite2[i]=sheet2.createRow((short)i);
				 fCell2 = row2.getFirstCellNum(); 
		         lCell2 = row2.getLastCellNum();	
		         for (int iCell = fCell2; iCell < lCell2; iCell++) {
					 cell2 = row2.getCell(iCell);
					 if(cell2==null){
						 continue;
					 }
					 else{
							//getting reference of current cell
							 Cell currentCell = cell2;
							 sheet2.autoSizeColumn(iCell);
							 DataFormatter dataFormatter = new DataFormatter();
							 String cellStringValue = dataFormatter.formatCellValue(row2.getCell(iCell));
							 rowwrite2[i].createCell(iCell+1).setCellValue(cellStringValue);
							 
							 if(i>=6 && iCell==10||i>=6 && iCell==12 ||i>=6 && iCell==13||i>=6 && iCell==14||i>=6 && iCell==20 ){
				    			 CellStyle dateStyle = wb.createCellStyle();
				    		       dateStyle.setDataFormat(
				    		           createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
				    		       Cell writeDate = rowwrite2[i].createCell(iCell+1);
				   	            writeDate.setCellValue(row2.getCell(iCell).getDateCellValue());
				   	            writeDate.setCellStyle(dateStyle); 
				   	         sheet2.setColumnWidth(iCell,1100*4);
				        continue;
							 }
							 
							 
							 
							 
							 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
			                     System.out.print(currentCell.getNumericCellValue() + "--");
			                     
			                     rowwrite2[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
			                     System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
			                     System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
			                    System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());    
							 	}
					 }
		     	}//for ends
			         	if(i>=6){
			           	 rowwrite2[i]=sheet2.getRow((short)i);;
			    		 rowwrite2[i].createCell(0).setCellFormula("CONCATENATE(H"+counter1+",U"+counter2+")");
			    		 CellReference cellReference = new CellReference("A"+counter1);
			    	 		Row rowF = sheet2.getRow(cellReference.getRow());
			    	 		Cell cellF = rowF.getCell(cellReference.getCol()); 
			    	 		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
			    	 		CellValue cellValue = evaluator.evaluate(cellF);
			    	 		System.out.println("  "+cellValue.getStringValue());
			    	 		rowwrite2[i].createCell(0).setCellValue(cellValue.getStringValue());
			    		 counter1+=1;counter2+=1;
			            }
			         	if(i==5){
			         		rowwrite[i]=sheet2.getRow((short)i);;
			         		 rowwrite[i].createCell(0).setCellValue("Validation Index");
			        	}
						}		

			            System.out.println("WorkBook has been created");
			}
	   
	   
	   
	  /* for(int i=rowStart2;i<=rowEnd2;i++){
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
	}*/
	   	//sheet2.createRow(0).createCell(0).setCellValue("ID");
		FileOutputStream fileOut = new FileOutputStream("Njoyn Master Tracker-16-18.xls(formatted).xls");
		 wb.write(fileOut);;
		fileOut.close();
		System.out.println("WorkBook has been created");
}

public static void main(String[] args) throws IOException{
	new SheetCopy();
}
}
