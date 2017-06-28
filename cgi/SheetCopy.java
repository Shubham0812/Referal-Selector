package cgi;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
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
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SheetCopy {

	
SheetCopy(String output1,String output2) throws IOException{
	
	
	
	 Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
	 Sheet sheet1 = wb.createSheet("Sheet1");
	 Sheet sheet2 = wb.createSheet("Sheet2");
	 int puu =  7;
     CellStyle style = wb.createCellStyle();
     Font font = wb.createFont();
     font.setFontHeightInPoints((short)11);
     font.setFontName(HSSFFont.FONT_ARIAL);
     font.setBold(true);
     style.setFont(font); 
	 
	 FileInputStream myStream = new FileInputStream(output1);
	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	 HSSFWorkbook wbread = new HSSFWorkbook(fs.getRoot(), true);
	 
	 FileInputStream myStream2 = new FileInputStream(output2);
	 NPOIFSFileSystem fs2 = new NPOIFSFileSystem(myStream2);
	 HSSFWorkbook wbread2 = new HSSFWorkbook(fs2.getRoot(), true);	
	 CreationHelper createHelper = wb.getCreationHelper();
	 HSSFSheet sheetx  = wbread.getSheetAt(0);
	 HSSFSheet sheetx2 = wbread2.getSheetAt(0);
	 HSSFRow row,row2;
	    int counter1 = 7;int counter1a = 7;
	    int counter2 = 7;int counter2a = 7;int count = 7;
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
				rowwrite[i]=sheet1.createRow((short)i);
				//first and last cell for the row
				 fCell = row.getFirstCellNum(); 
		         lCell = row.getLastCellNum();	
		         for(int iCell = fCell; iCell < lCell; iCell++) {
		         cell = row.getCell(iCell);
				 if(cell==null){
					 if(iCell==9){
						 Cell currentCells = row.getCell(iCell+3);
						 if(currentCells==null){
							 Cell currentCeller = row.getCell(iCell+4);
							 if(currentCeller.getCellTypeEnum() == CellType.NUMERIC){
								 double value = currentCeller.getNumericCellValue();
								 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+value+",10)");
								 continue;}
				    		     else if(currentCeller.getCellTypeEnum() == CellType.STRING){
				    		    	 String value = currentCeller.getStringCellValue();
			    					 try{
			    						 String newValue = value.replaceAll("-","");
			    						 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
			    						 continue;
			    					 	}catch(Exception e){
			    						 String newValue = value.replaceAll("\\s","");
			        					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");}
			    					 	continue;}}
		    				   	 if(currentCells.getCellTypeEnum() == CellType.NUMERIC){
		    				   		 double value = currentCells.getNumericCellValue();
		    				  		 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+value+",10)");
		    				  		 continue;}
		    				  	else if(currentCells.getCellTypeEnum() == CellType.STRING){
		    				  		String value = currentCells.getStringCellValue();
		    				  		try{
		    				  			String newValue = value.replaceAll("-","");
		    				  			rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
		    				  			continue;
		    				  			}catch(Exception e){
		    						 String newValue = value.replaceAll("\\s","");
		        					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
		        					 continue;
		    					 }}}}
				 //if the cell has value determine the type of value.
				 else{
				 //getting reference of current cell
					 Cell currentCell = cell;
					 sheet1.autoSizeColumn(iCell);
				 // testing for types of the cell
					 DataFormatter dataFormatter = new DataFormatter();
					// String cellStringValue = dataFormatter.formatCellValue(row.getCell(iCell));
				   	// rowwrite[i].createCell(iCell+1).setCellValue(cellStringValue);	 
					 if(i>=6 && iCell==5 ||i>=6 && iCell==6){
						 CellStyle dateStyle = wb.createCellStyle();
			    		 dateStyle.setDataFormat(
			    		 createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			    		 Cell writeDate = rowwrite[i].createCell(iCell+1);
			   	         writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
			   	         writeDate.setCellStyle(dateStyle); 
			   	         sheet1.setColumnWidth(iCell,1100*4);
			   	         sheet1.setColumnWidth(20,1200*4);
			   	         sheet1.setColumnWidth(21,1200*4);
			   	         sheet1.setColumnWidth(22,1400*4);
			   	         sheet1.setColumnWidth(23,1400*4);
			   	         continue;}
					 
					 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
				//		 System.out.print(currentCell.getNumericCellValue() + "--");       
		    			 if(i>=6&& iCell ==9){
		    				 double value = currentCell.getNumericCellValue();
		    				 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+value+",10)");
		    				 continue;}
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
						 }
					 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
				//		 System.out.print(currentCell.getStringCellValue() + "--");
		    			 if(i>=6&& iCell ==9){
		    				 String value = currentCell.getStringCellValue();
		    				 try{
		    					 String newValue = value.replaceAll("-","");
		    					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
		    					 }catch(Exception e){
		    					 String newValue = value.replaceAll("\\s","");
		        				 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");}
		    					 continue;}
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());}
						 else if(currentCell.getCellTypeEnum() == CellType.FORMULA){
							 System.out.print(currentCell.getStringCellValue() + "--");
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());}
						 else if (currentCell.getCellTypeEnum() == CellType.ERROR){
		                  // System.out.print(currentCell.getStringCellValue() + "--");
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());}
					 
					 
					 
}}//cell for loop ends
		          
		         //Validation Index Calculation
		         	if(i>=6){
		         		rowwrite[i]=sheet1.getRow((short)i);;
		         		rowwrite[i].createCell(0).setCellFormula("CONCATENATE(F"+counter1+",D"+counter2+")");
		         		CellReference cellReference = new CellReference("A"+counter1);
		         		Row rowF = sheet1.getRow(cellReference.getRow());
		         		Cell cellF = rowF.getCell(cellReference.getCol()); 
		         		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
		         		CellValue cellValue = evaluator.evaluate(cellF);
		         		System.out.println("  "+cellValue.getStringValue());
		         		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
		         		counter1+=1;counter2+=1;}

					if(i>=6){
						 rowwrite[5].createCell(20).setCellValue("Referred By");
						 rowwrite[i].createCell(20).setCellFormula("VLOOKUP(A"+count+",Sheet2!1:65536,6,0)");				
						 
						 
						 rowwrite[5].createCell(21).setCellValue("Referred By Email");
						 rowwrite[i].createCell(21).setCellFormula("VLOOKUP(A"+count+",Sheet2!1:65536,7,0)");
					 
						 rowwrite[5].createCell(22).setCellValue("Referred By");
						 rowwrite[i].createCell(22).setCellFormula("IF(T"+count+"=0,U"+count+",T"+count+")");
						 
			         		
						 rowwrite[5].createCell(23).setCellValue("Referred By Email");
						 rowwrite[i].createCell(23).setCellFormula("IF(S"+count+"=0,V"+count+",S"+count+")");
						 count++;
						 
					}
		         	
		         	
		         	if(i==5){
		         		rowwrite[i]=sheet1.getRow((short)i);;
		         		rowwrite[i].createCell(0).setCellValue("Validation Index");}
}//row not null ends

		      System.out.println("WorkBook has been created");
		      if(i>=5){		    
		    	  for(int x =0;x<rowwrite[5].getLastCellNum();x++){
			   Cell co = sheet1.getRow(5).getCell(x);
				co.setCellStyle(style);
		      }}
		      

}
	
	   
	   /*for(int i=rowStart;i<=rowEnd;i++){
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
	   */
		for(int i1=rowStart2;i1<=rowEnd2;i1++){
			row2=sheetx2.getRow(i1);
			if(row2==null){
				System.out.println("empty accessed");
				continue;
				}
			if(row2!=null){
				rowwrite2[i1]=sheet2.createRow((short)i1);
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
							 rowwrite2[i1].createCell(iCell+1).setCellValue(cellStringValue);
							 
							 if(i1>=6 && iCell==10||i1>=6 && iCell==12 ||i1>=6 && iCell==13||i1>=6 && iCell==14||i1>=6 && iCell==20 ){
				    			 CellStyle dateStyle = wb.createCellStyle();
				    		       dateStyle.setDataFormat(
				    		           createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
				    		       Cell writeDate = rowwrite2[i1].createCell(iCell+1);
				   	            writeDate.setCellValue(row2.getCell(iCell).getDateCellValue());
				   	            writeDate.setCellStyle(dateStyle); 
				   	         sheet2.setColumnWidth(iCell,1100*4);
				        continue;
							 }
							 
							 
							 
							 
							 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
			                     System.out.print(currentCell.getNumericCellValue() + "--");
			                     
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
			                     System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
			                     System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getCellFormula());    
							 }
							 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
			                    System.out.print(currentCell.getStringCellValue() + "--");
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());    
							 	}
					 }
		     	}//for ends
			         	if(i1>=6){
			           	 rowwrite2[i1]=sheet2.getRow((short)i1);;
			    		 rowwrite2[i1].createCell(0).setCellFormula("CONCATENATE(H"+counter1a+",U"+counter2a+")");
			    		 CellReference cellReference = new CellReference("A"+counter1a);
			    	 		Row rowF = sheet2.getRow(cellReference.getRow());
			    	 		Cell cellF = rowF.getCell(cellReference.getCol()); 
			    	 		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
			    	 		CellValue cellValue = evaluator.evaluate(cellF);
			    	 		System.out.println("  "+cellValue.getStringValue());
			    	 		rowwrite2[i1].createCell(0).setCellValue(cellValue.getStringValue());
			    		 counter1a+=1;counter2a+=1;
			            }
			         	if(i1==5){
			         		rowwrite2[i1]=sheet2.getRow((short)i1);;
			         		 rowwrite2[i1].createCell(0).setCellValue("Validation Index");
			        	}
						}		

			            System.out.println("WorkBook has been created");
		      if(i1>=5){		    
		    	  for(int x =0;x<rowwrite2[5].getLastCellNum();x++){
			   Cell co = sheet2.getRow(5).getCell(x);
				co.setCellStyle(style);
		      }}			            
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
		FileOutputStream fileOut = new FileOutputStream("VLookupFile.xls");
		 wb.write(fileOut);;
		fileOut.close();
		System.out.println("Finale WorkBook has been created"); 
		File output = new File("VLookupFile.xls");
		String path = output.getPath();
		Runtime.getRuntime().exec("explorer.exe /select," + path);
}
			

public static void main(String[] args) throws IOException{
	String output1 = "Njoyn Master Tracker-16-18.xls";
	String output2 ="Candidate Referrals (Generic).xls";
	new SheetCopy(output1,output2);
}
}
