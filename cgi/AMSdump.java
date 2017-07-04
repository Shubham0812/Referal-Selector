//********************************************
//code to write the contents of two different Excel files onto a single file with 2 sheets to apply formulas between them.
//Author - Shubham kr. Singh
//*********************************************

package cgi;
import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;

import cgi.AMSdump.DataStorer;
public class AMSdump {
	
	
	java.util.List<DataStorer> data = new ArrayList<DataStorer>();
	
	
	static String open = "";
	static String output1;
	static String output2;
	 static XSSFWorkbook wb = new XSSFWorkbook();
	 Sheet sheet1 = wb.createSheet("Sheet1");
	 Sheet sheet2 =wb.createSheet("Sheet2");
	 int LastRowNum;
	 int morePasses;
	 FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
//constructor of the class
AMSdump(String output1,String output2) throws IOException,InvalidFormatException{
	
//CellStyle to set the cell headers as bold
	
	 CellStyle style = wb.createCellStyle();
	 Font font = wb.createFont();
     font.setFontHeightInPoints((short)11);
     font.setFontName(HSSFFont.FONT_ARIAL);
     font.setBold(true);
     style.setFont(font); 
     
//to open first file
     
     FileInputStream myStream = new FileInputStream(output1);
	 OPCPackage pkg = OPCPackage.open(myStream);
	 XSSFWorkbook wbread = new XSSFWorkbook(pkg);
//creating helper to edit and apply formulas to the output workbook
	 
	 CreationHelper createHelper = wb.getCreationHelper();
	 XSSFSheet sheetx  = wbread.getSheetAt(0);
	 XSSFRow row;
	 XSSFCell cell;
//getting the first and last row of the input Sheet 1
	 
	 int rowStart = sheetx.getFirstRowNum() ;
     int rowEnd = sheetx.getLastRowNum() ;int count = 7;
	 int fCell,lCell;
	 Row rowwrite[] =new Row[rowEnd+1];
	
//writing the contents of first file to AMS_DUMP_OUTPUT.XLSX";
	 
		DumpWrite(output2);
	 
	 for(int i=rowStart;i<=rowEnd;i++){
		 row=sheetx.getRow(i);
		 if(row==null){
		 System.out.println("empty accessed");
		 continue;}
		 if(row!=null){
		 rowwrite[i]=sheet1.createRow((short)i);
						
//to get the first and last cell of the row
		 
		 fCell = row.getFirstCellNum(); 
		 lCell = row.getLastCellNum();	
		 
//iterating over the cells of a particular row and writing it one by one in the workbook
		 
		 for(int iCell = fCell; iCell < lCell; iCell++) {
			 cell = row.getCell(iCell);
			 
//if cell has no value do nothing skip it and continue
			 
			 if(cell==null){
				 continue;}
			 
//if the cell has value determine the type of value.
			 
			 else{
				 
//getting reference of current cell
				 
				 	Cell currentCell = cell;
				 	
//reading and writing dates require special data format
				 	
				 	if(i>=6 && iCell==6 ||i>=6 && iCell==7){
				 	try {
				 			CellStyle dateStyle = wb.createCellStyle();
				 			dateStyle.setDataFormat(
				 			createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			    		    Cell writeDate = rowwrite[i].createCell(iCell);
			   	            writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
			   	            writeDate.setCellStyle(dateStyle); 
			   	            continue;
			   	           }catch(Exception ex){}
				 	
//setting the Width of the column with dates
				 	
						    sheet1.setColumnWidth(iCell,1300*4);
						    }
				 	
//setting the Width of 13th column
				 	
					 if(i==12){
						 sheet1.setColumnWidth(12,1800*4);
					 }
					 
//determining the type of cell being read and writing that to the new workbook
					 
					 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) { 
		                   rowwrite[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue());}
					 else if(currentCell.getCellTypeEnum() == CellType.STRING) {
		                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getStringCellValue());}
					 else if(currentCell.getCellTypeEnum() == CellType.FORMULA){
		                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getCellFormula());}
					 else if (currentCell.getCellTypeEnum() == CellType.ERROR){
		                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getErrorCellValue());}
			 	}//else part of cell ends
			 
			 }//inner cell for loop ends
		 

			if(i>=6){
				 rowwrite[5].createCell(24).setCellValue("Mobile Check");
				 rowwrite[i].createCell(24).setCellFormula("VLOOKUP(K"+count+",Sheet2!A:A,1,0)");			
				 
				 rowwrite[5].createCell(25).setCellValue("Email Check");
				 rowwrite[i].createCell(25).setCellFormula("VLOOKUP(E"+count+",Sheet2!B:B,1,0)");
	 			/*	CellReference cellReference = new CellReference("Z"+(i+1));
					 Row rowF = sheet1.getRow(cellReference.getRow());
		        		Cell cellF = rowF.getCell(cellReference.getCol());
		        		CellValue cellValue = evaluator.evaluate(cellF);
         				System.out.println("  "+cellValue.getStringValue());
		        		rowwrite[i].createCell(25).setCellValue(cellValue.getStringValue()); */
					 count++;
			}
		}//row not null ends
		 System.out.println("WorkBook has been created");
		 
//setting the font as bold for the headers
		 

		 
}//outer for loop ends

//Writing to the file
	 	FileOutputStream fileOut = new FileOutputStream("AmsDumpOutput"+open+".xlsx");
	    wb.write(fileOut);
		fileOut.close();
	    wbread.close();
		System.out.println("Sheet1 of WorkBook has been created");
		
//calling the function to read data from Second file and writing it to the Sheet2 of the same output Workbook
		
	
	
//Getting the number of rows still remaining to be read
	
	System.out.println("Remaining are :- " + (LastRowNum-59805));
	
//Passes required to fully read the second file
	
	morePasses = (LastRowNum/59805);
	System.out.println("Passes Required are : "+(morePasses));
	
	for(int i = 0;i<20000;i++){
		System.out.println("EMAIL : " +data.get(i).getEmail()+ "mobile : " +data.get(i).getMobile());
	}
	System.out.println(data.get(1).getEmail());

}

//////////////////////////////////////////////////////////////////////////////////////////////////////
//function to read the Second file and write to the Same output Workbook :- "AMS_DUMP_OUTPUT.XLSX"

public void DumpWrite(String output2) throws IOException{
	
//opening the second file to read from using Streaming Reader
	
	File is = new File(output2);
	Workbook workbook = StreamingReader.builder()
						.rowCacheSize(100)
						.bufferSize(9096)
						.open(is);
	int cols = 0;int rows = 0;
	Sheet  sheet22 = workbook.getSheetAt(0);
	
//getting the number of rows in the second input file
	
	LastRowNum = sheet22.getLastRowNum();
    int rowCounter = 1;
    int perRow = 0;int count = 0;
//creating an array of rows to accommodate the number of rows in the input file
    
    Row rowwrite[] =new Row[LastRowNum+1];
    
//Iterator to iterate the second input file
		
    for (Sheet sheet : workbook){
    	
//getting the Sheet name and the lastRow of the Input Sheet
    	
	    System.out.println(sheet.getSheetName() + " Rows :  "+sheet.getLastRowNum());
    	for (Row r : sheet) {
    		
//creating a new row in sheet2
    		DataStorer obj = new DataStorer();
    		rowwrite[rows] = sheet2.createRow((int)rows);
    		for (Cell c : r) {
    			
    			System.out.println(c.getRowIndex());
//if the content of the second input sheet is number write number to the second sheet of output file   			

    			if(c.getCellTypeEnum()==CellType.NUMERIC){
    				if(cols==0){
    				obj.setMobile(c.getNumericCellValue());
    				rowwrite[rows].createCell(cols).setCellValue(c.getNumericCellValue());
    				cols++;count++;
    				continue;}
    				if(cols==2){
    					obj.setID(c.getNumericCellValue());
    				rowwrite[rows].createCell(cols).setCellValue(c.getNumericCellValue());
    				cols++;count++;
    				continue;}
    			}
    			
//if the content of the second input sheet is String write that String to the output workbook
    			if(cols==1){
    				obj.setEmail(c.getStringCellValue());
    			rowwrite[rows].createCell(cols).setCellValue(c.getStringCellValue());
	    		cols++;}
    			if(cols==3){
    				obj.setPan(c.getStringCellValue());
        			rowwrite[rows].createCell(cols).setCellValue(c.getStringCellValue());
    	    		cols++;}
	    		
//the heap goes out of memory and it shows that it cannot process more than 59010 rows of the second sheet which as of row contains 161201 records
//using if to end The rows at 59010 and create the final output files.
	    		
	    		if(rowCounter ==56000)
	    		{
	    			FileOutputStream fileOuts = new FileOutputStream("AmsDumpOutput"+open+".xlsx");
	    			wb.write(fileOuts);;
	    			fileOuts.close();
	    			
//closing the output workbook
	    		
	    			System.out.println("Sheet 2 of Output WorkBook has been created");
	    			
//using runtime to open the file where the output was created
	    			
	    			File output = new File("AmsDumpOutput"+open+".xlsx");
	    			String path = output.getPath();
	    			Runtime.getRuntime().exec("explorer.exe /select," + path);
	    			return;	
	    		}
    						}
    		data.add(obj);
	      rows++;perRow++;//iterating the row to next row
	      
	      cols=0;count=0;//setting the columns again to zero
	      
	      rowCounter++;
	    				}
	}
}

public class DataStorer{
	double mobile;
	String email;
	double ID;
	String Pan;
	
	public void setMobile(double mobile){
		this.mobile = mobile;
	}
	
	public double getMobile(){
		return mobile;
	}
	public void setEmail(String email){
		this.email = email;
	}
	
	public String getEmail(){
		return email;
	}
	public void setID(double ID){
		this.ID = ID;
	}
	
	public double getID(){
		return ID;
	}
	public void setPan(String Pan){
		this.Pan = Pan;
	}
	
	public String getPan(){
		return Pan;
	}
	

	}




//main function to call the constructor

	public static void main(String[] args) throws IOException,InvalidFormatException{
		output1 = "VLookupOutputs.xlsx";	//first input file to be passed
		output2 ="AMS_Dump_Data"+open+"x.xlsx";  //second input file to be passed
		new AMSdump(output1,output2); 		   //calling the constructor
	}
}
