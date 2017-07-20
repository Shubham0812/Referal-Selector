/* Automation of Member Referral Process (June-2017) 
 * Author - Shubham Kumar Singh
 * Email - singh.shubham0812@gmail.com
 * College - Nitte Meenakshi Institute of Technology, Bangalore 
 */

//this code copies both Master Tracker & Candidate Referral into 2 Sheets of output file and Performs Vlookup as well as formatting
package cgi;

import java.awt.Toolkit;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetCopy {

	SheetCopy(String output1,String output2) throws IOException{
	
	int date1=0,date2=0;
	int phone = 0;String phoneA = "";
	int req = 0;String reqA = "";
	int can = 0;String canA = "";
	int cellPhone = 0;
	int date1x=0,date2x=0,date3=0,date4=0,date5=0;
	int job = 0;String jobA = "";
	int canx = 0;String canxA = "";
	int ref = 0; int ref2 = 0;	
	Workbook wb = new XSSFWorkbook();  // or new XSSFWorkbook();
	Sheet sheet1 = wb.createSheet("Sheet1");
 	Sheet sheet2 = wb.createSheet("Sheet2");
 	CellStyle style = wb.createCellStyle();
 	Font font = wb.createFont();
 	font.setFontHeightInPoints((short)11);
 	font.setFontName(HSSFFont.FONT_ARIAL);
 	font.setBold(true);
 	style.setFont(font); 
 	String number_c;
 	CellStyle num = wb.createCellStyle();
 	num.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
 	NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(output1));
 	HSSFWorkbook wbread = new HSSFWorkbook(fs.getRoot(), true);
 	NPOIFSFileSystem fs2 = new NPOIFSFileSystem(new File(output2));
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
    try{ 
    	List<String> headings = new ArrayList<String>();
	    Row extra = sheetx.getRow(5);
	    for(int counter=0;counter<extra.getLastCellNum();counter++){
		    Cell extraCell = extra.getCell(counter);
		    headings.add(extraCell.getStringCellValue());}
	    for(int ca = 0;ca<headings.size();ca++){
	    	if(headings.get(ca).toString().equals("Applied Date (WEB)")){
	    		date1=ca;}
	    	if(headings.get(ca).toString().equals("Applied Date (WEB/MCH)")){
	    		date2=ca;}
	    	if(headings.get(ca).toString().equals("Candidate Phone Number")){
	    		phone=ca;
	    		phoneA = Intro.checkAlphabet(phone+1);}
	    	if(headings.get(ca).toString().equals("Cell Phone")){
	    		cellPhone=ca;}
	    	if(headings.get(ca).toString().equals("REQ #")){
	    		req=ca;
	    		reqA = Intro.checkAlphabet(req+1);}
	    	if(headings.get(ca).toString().equals("Candidate ID")){
	    		can=ca;
	    		canA = Intro.checkAlphabet(can+1);}
		    }
		    }catch(NullPointerException e){}
		try{ 
	    	 List<String> headinge = new ArrayList<String>();
	    	 Row extra = sheetx2.getRow(5);
	    	 for(int counter=0;counter<extra.getLastCellNum();counter++){
	    		 Cell extraCell = extra.getCell(counter);
	    		 headinge.add(extraCell.getStringCellValue());}
	    	 for(int ca = 0;ca<headinge.size();ca++){
	    	   	if(headinge.get(ca).toString().equals("Referral Name")){
	    	   		ref=ca+2;}
	    	    if(headinge.get(ca).toString().equals("Referral Email")){
	    	    		ref2=ca+2;}	
	    	    	
	    	    }
	    	    }catch(NullPointerException e){}	   
	   
	   
		for(int i=rowStart;i<=rowEnd;i++){
			row=sheetx.getRow(i);
			if(row==null){continue;}
			if(row!=null){
				rowwrite[i]=sheet1.createRow((short)i);
				//first and last cell for the row
				fCell = row.getFirstCellNum(); 
		        lCell = row.getLastCellNum();	
		        for(int iCell = fCell; iCell < lCell; iCell++) {
		        	cell = row.getCell(iCell);
		        	if(cell==null){continue;}
				//if the cell has value determine the type of value.
				 else{
				 //getting reference of current cell
					Cell currentCell = cell;
					sheet1.autoSizeColumn(iCell);				
					if(i>=6 && iCell==date1 ||i>=6 && iCell==date2){
						 try{
						 CellStyle dateStyle = wb.createCellStyle();
			    		 dateStyle.setDataFormat(
			    		 createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			    		 Cell writeDate = rowwrite[i].createCell(iCell+1);
			   	         writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
			   	         writeDate.setCellStyle(dateStyle); 
			   	         sheet1.setColumnWidth(iCell,1100*4); sheet1.setColumnWidth(20,1200*4);
			   	         sheet1.setColumnWidth(21,1200*4); sheet1.setColumnWidth(22,1400*4); sheet1.setColumnWidth(23,1400*4);
			   	         continue;
			   	         }catch(Exception ex){}
					 }
					 if(i==12){
						 sheet1.setColumnWidth(12,1800*4);
					 }
					 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {     
		    			 if(i>=6&& iCell ==phone){
		    				 double value = currentCell.getNumericCellValue();
		    				 String axe =""+currentCell.getAddress();
		    				 if(axe.length()==2){number_c = axe.substring(1,2);}
		    				 else if(axe.length()==3){number_c = axe.substring(1,3);}
		    				 else if(axe.length()==4){number_c=axe.substring(1,4);}
		    				 else{number_c=axe.substring(1,5);}
		    				 rowwrite[i] = sheet1.getRow((short)i);
		    				 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+value+",10)");
		    				 CellReference cellReference = new CellReference(phoneA+number_c);
		    				 Row rowF = sheet1.getRow(cellReference.getRow());
		    				 Cell cellF = rowF.getCell(cellReference.getCol()); 
     						 CellValue cellValue = evaluator.evaluate(cellF);
     						 Cell xcu =rowwrite[i].createCell(iCell+1);
		    	         	 xcu.setCellStyle(num);
		    	         	 xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		    				 continue; }
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
						 }
					 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
		    			 if(i>=6&& iCell ==phone){
		    				 try{
		    					 Row are = sheetx.getRow(i);
		    					 if(are.getCell(phone).getStringCellValue().equals(" ")){
		    						 try{
	    								 Cell currentCells = row.getCell(cellPhone);
	    								 if (currentCells.getCellTypeEnum() == CellType.NUMERIC) {
	    									 double value = currentCells.getNumericCellValue();
		    			    				 String axe =""+currentCells.getAddress();
		    			    				 if(axe.length()==2){number_c = axe.substring(1,2);}
		    			    				 else if(axe.length()==3){number_c = axe.substring(1,3);}
		    			    				 else if(axe.length()==4){number_c=axe.substring(1,4); }
		    			    				 else{number_c=axe.substring(1,5); }
		    			    				 rowwrite[i] = sheet1.getRow((short)i);
		    			    				 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+value+",10)");
		    			    				 CellReference cellReference = new CellReference(phoneA+number_c);
		    			    				 Row rowF = sheet1.getRow(cellReference.getRow());
    			    	         			 Cell cellF = rowF.getCell(cellReference.getCol()); 
		    			    	         	 CellValue cellValue = evaluator.evaluate(cellF);
		    			    	         	 Cell xcu =rowwrite[i].createCell(iCell+1);
		    			    	         	 xcu.setCellStyle(num);
		    				         		 xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		    			    				 continue;
		    								 }
		    								 else if (currentCells.getCellTypeEnum() == CellType.STRING) {
		    				    				 String add =""+currentCell.getAddress();
		    				    				 if(add.length()==2){number_c = add.substring(1,2);}
		    				    				 else if(add.length()==3){number_c = add.substring(1,3);}
		    				    				 else if(add.length()==4){number_c=add.substring(1,4); }
		    				    				 else{number_c=add.substring(1,5); }
		    				    				 String value = currentCells.getStringCellValue();
		    				    				 try{
		    				    					 String newValue = value.replaceAll("-","");
		    				    					 rowwrite[i] = sheet1.getRow((short)i);
		    				    					 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+newValue+",10)");
		    				    					 CellReference cellReference = new CellReference(phoneA+number_c);
		    				        				 Row rowF = sheet1.getRow(cellReference.getRow());
		    				        	         	 Cell cellF = rowF.getCell(cellReference.getCol()); 
		    				        	         	 CellValue cellValue = evaluator.evaluate(cellF);
		    				        	         	 Cell xcu =rowwrite[i].createCell(iCell+1);
		    				            	         xcu.setCellStyle(num);
	    				            	         	 xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		    				            	         continue;
		    				    					 }catch(Exception e){
		    				    					 String newValue = value.replaceAll("\\s","");
		    				    					 try{
		    				    						 rowwrite[i] = sheet1.getRow((short)i);
		    				        					 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+newValue+",10)");
		    				        					 CellReference cellReference = new CellReference(phoneA+number_c);
		    				            				 Row rowF = sheet1.getRow(cellReference.getRow());
		    				            	         	 Cell cellF = rowF.getCell(cellReference.getCol()); 
		    				            	         	 CellValue cellValue = evaluator.evaluate(cellF);
		    				            	         	 Cell xcu =rowwrite[i].createCell(iCell+1);
		    				                	         xcu.setCellStyle(num);
		    				                	         xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		    				        				 }catch(Exception af){}
		    				    					 }
		    				    					 continue;
		    								 }
		    								 
		    							 }catch(NullPointerException nula){} 
		    					 }
		    				 }catch(NullPointerException a){}
		    				 String add =""+currentCell.getAddress();
		    				 if(add.length()==2){number_c = add.substring(1,2);}
		    				 else if(add.length()==3){number_c = add.substring(1,3);}
		    				 else if(add.length()==4){number_c=add.substring(1,4);}
		    				 else{number_c=add.substring(1,5); }
		    				 String value = currentCell.getStringCellValue();
		    				 try{
		    					 String newValue = value.replaceAll("-","");
		    					 rowwrite[i] = sheet1.getRow((short)i);
		    					 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+newValue+",10)");
		    					 CellReference cellReference = new CellReference(phoneA+number_c);
		        				 Row rowF = sheet1.getRow(cellReference.getRow());
	        	         		 Cell cellF = rowF.getCell(cellReference.getCol()); 
	        	         		 CellValue cellValue = evaluator.evaluate(cellF);
		        	         	 Cell xcu =rowwrite[i].createCell(iCell+1);
		            	         xcu.setCellStyle(num);
		            	         xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		            	         continue;
		    					 }catch(Exception e){
		    					 String newValue = value.replaceAll("\\s","");
		    					 try{
		    						 rowwrite[i] = sheet1.getRow((short)i);
		    						 rowwrite[i].createCell(phone+1).setCellFormula("RIGHT("+newValue+",10)");
		    						 CellReference cellReference = new CellReference(phoneA+number_c);
		            				 Row rowF = sheet1.getRow(cellReference.getRow());
		            	         	 Cell cellF = rowF.getCell(cellReference.getCol()); 
		            	         	 CellValue cellValue = evaluator.evaluate(cellF);
		            	         	 Cell xcu =rowwrite[i].createCell(iCell+1);
		                	         xcu.setCellStyle(num);
		                	         xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
		        				 }catch(Exception af){}
		    					 }
		    					 continue;}
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());}
						 else if(currentCell.getCellTypeEnum() == CellType.FORMULA){
							 System.out.println(currentCell.getStringCellValue() + "--");
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());}
						 else if (currentCell.getCellTypeEnum() == CellType.ERROR){
		                   System.out.println(currentCell.getStringCellValue() + "--");
		                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());}
					 
					 
					 
}}//cell for loop ends
		          
		         //Validation Index Calculation
		         	if(i>=6){
		         		rowwrite[i]=sheet1.getRow((short)i);;
		         		rowwrite[i].createCell(0).setCellFormula("CONCATENATE("+reqA+counter1+","+canA+counter2+")");
		         		CellReference cellReference = new CellReference("A"+counter1);
		         		Row rowF = sheet1.getRow(cellReference.getRow());
		         		Cell cellF = rowF.getCell(cellReference.getCol()); 
		         		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
		         		CellValue cellValue = evaluator.evaluate(cellF);
		         		//System.out.println("  "+cellValue.getStringValue());
		         		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
		         		counter1+=1;counter2+=1;}

					if(i>=6){
						 rowwrite[5].createCell(20).setCellValue("Referred By");
						 rowwrite[i].createCell(20).setCellFormula("VLOOKUP(A"+count+",Sheet2!1:65536,"+ref+",0)");				
						 rowwrite[5].createCell(21).setCellValue("Referred By Email");
						 rowwrite[i].createCell(21).setCellFormula("VLOOKUP(A"+count+",Sheet2!1:65536,"+ref2+",0)");
						 count++;}
					if(i==5){
		         		rowwrite[i]=sheet1.getRow((short)i);;
		         		rowwrite[i].createCell(0).setCellValue("Validation Index");}
}//row not null ends

			  //setting bold on the column headers
		      if(i>=5){
		    	  try{
		    	  for(int x =0;x<rowwrite[5].getLastCellNum();x++){
			   Cell co = sheet1.getRow(5).getCell(x);
				co.setCellStyle(style);
		    	  }
		      }catch(NullPointerException er){continue;}
		    	  }
		      
//outer foor loop for sheet 1 writing ends
}
		
		try{ 
	    	 List<String> heading = new ArrayList<String>();
	    	 Row extra = sheetx2.getRow(5);
	    	    for(int counter=0;counter<extra.getLastCellNum();counter++){
	    	        Cell extraCell = extra.getCell(counter);
	    	        heading.add(extraCell.getStringCellValue());
	    	        }
	    	    for(int ca = 0;ca<heading.size();ca++){
	    	    	
	    	    	if(heading.get(ca).toString().equals("Application Date")){
    	    		date1x=ca;}
	    	    	if(heading.get(ca).toString().equals("Date Survey Taken")){
	    	    		date2x=ca;}
	    	    	if(heading.get(ca).toString().equals("Date Survey Invite Sent")){
	    	    		date3=ca;}
	    	    	if(heading.get(ca).toString().equals("Candidate Enter Date")){
	    	    		date4=ca;}
	    	    	if(heading.get(ca).toString().equals("Last Activity Date")){
	    	    		date5=ca;}
	    	    	if(heading.get(ca).toString().equals("Job ID")){
	    	    		job=ca;
	    	    		jobA = Intro.checkAlphabet(job+1);}
	    	    	if(heading.get(ca).toString().equals("CandidateID")){
	    	    		canx=ca;
	    	    		canxA = Intro.checkAlphabet(canx+1);}
	    	    	}
	    	    }catch(NullPointerException e){}
		
		for(int i1=rowStart2;i1<=rowEnd2;i1++){
			row2=sheetx2.getRow(i1);
			if(row2==null){continue;}
			if(row2!=null){
				rowwrite2[i1]=sheet2.createRow((short)i1);
				 fCell2 = row2.getFirstCellNum(); 
		         lCell2 = row2.getLastCellNum();	
		         for (int iCell = fCell2; iCell < lCell2; iCell++) {
					 cell2 = row2.getCell(iCell);
					 if(cell2==null){continue;}
					 else{
							//getting reference of current cell
							 Cell currentCell = cell2;
							 sheet2.autoSizeColumn(iCell);			 
							 if(i1>=6 && iCell==date1x||i1>=6 && iCell==date2x ||i1>=6 && iCell==date3||i1>=6 && iCell==date4||i1>=6 && iCell==date5 ){
								 try{
									 CellStyle dateStyle = wb.createCellStyle();
									 dateStyle.setDataFormat(
				    		         createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
									 Cell writeDate = rowwrite2[i1].createCell(iCell+1);
									 writeDate.setCellValue(row2.getCell(iCell).getDateCellValue());
									 writeDate.setCellStyle(dateStyle); 
									 sheet2.setColumnWidth(iCell,1100*4);
									 continue;}catch(Exception e){}
							 }
							 
							 	if (currentCell.getCellTypeEnum() == CellType.NUMERIC){
							 		rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());}
							  else if (currentCell.getCellTypeEnum() == CellType.STRING){
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());}
							  else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getCellFormula());}
							 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
			                     rowwrite2[i1].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());    }
					 }
		     	}//for ends
			         	if(i1>=6){
			         		rowwrite2[i1]=sheet2.getRow((short)i1);;
			         		rowwrite2[i1].createCell(0).setCellFormula("CONCATENATE("+jobA+counter1a+","+canxA+counter2a+")");
			         		CellReference cellReference = new CellReference("A"+counter1a);
			    	 		Row rowF = sheet2.getRow(cellReference.getRow());
			    	 		Cell cellF = rowF.getCell(cellReference.getCol()); 
			    	 		CellValue cellValue = evaluator.evaluate(cellF);
			    	 		rowwrite2[i1].createCell(0).setCellValue(cellValue.getStringValue());
			    	 		counter1a+=1;counter2a+=1;
			            }
			         	if(i1==5){
			         		rowwrite2[i1]=sheet2.getRow((short)i1);;
			         		rowwrite2[i1].createCell(0).setCellValue("Validation Index");
			        	}
						}		
		      if(i1>=5){		    
		    	  for(int x =0;x<rowwrite2[5].getLastCellNum();x++){
		    		  Cell co = sheet2.getRow(5).getCell(x);
		    		  co.setCellStyle(style);
		      }}			            
}
	
		 //removing reference of the vLookup
		for (int x = 6;x<sheet1.getLastRowNum()+1;x++){
			
			 CellReference cellReference = new CellReference("U"+(x+1));
			 Row rowF = sheet1.getRow(cellReference.getRow());
         	 Cell cellF = rowF.getCell(cellReference.getCol());
			 CellValue cellValue = evaluator.evaluate(cellF);
     		 rowwrite[x].createCell(20).setCellValue(cellValue.getStringValue());
		}
		for (int x = 6;x<sheet1.getLastRowNum()+1;x++){
			CellReference cellReference = new CellReference("V"+(x+1));
			Row rowF = sheet1.getRow(cellReference.getRow());
    		Cell cellF = rowF.getCell(cellReference.getCol());
			CellValue cellValue = evaluator.evaluate(cellF);
    		rowwrite[x].createCell(21).setCellValue(cellValue.getStringValue());
		}
		//reference remover done
		FileOutputStream fileOut = new FileOutputStream("VLookupOutputs.xlsx");
		wb.write(fileOut);;
		fileOut.close();
		wb.close();wbread.close();fs.close();wbread2.close();fs2.close();
		System.out.println("Finale WorkBook has been created"); 
		Toolkit.getDefaultToolkit().beep();

}
			
//the main function of SheetCopy
public static void main(String[] args) throws IOException{
	String output1 = "(CGI) Requisition Applicants (5).xls";
	String output2 ="Copy of Candidate Referrals (Generic).xls";
	new SheetCopy(output1,output2);
}
}
