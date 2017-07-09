//********************************************
//code to write the contents of two different Excel files onto a single file with 2 sheets to apply formulas between them.
//Author - Shubham kr. Singh
//*********************************************

package cgi;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
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
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;
public class AMSdump {
	
	
	java.util.List<DataStorer> data = new ArrayList<DataStorer>();
	
	 int check;
	static String open = "";
	static String output1;
	static String output2;	    	  int check1,check2;
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
     
     Font font2 = wb.createFont();
     font2.setColor(IndexedColors.WHITE.getIndex());
     
     Font font3 = wb.createFont();
     font3.setColor(IndexedColors.BLACK.getIndex());
     
     CellStyle green = wb.createCellStyle();
     green.setFillForegroundColor(IndexedColors.GREEN.getIndex());
     green.setFillPattern(CellStyle.SOLID_FOREGROUND);
     green.setFont(font2);
     
     CellStyle red = wb.createCellStyle();
     red.setFillForegroundColor(IndexedColors.RED.getIndex());
     red.setFillPattern(CellStyle.SOLID_FOREGROUND);
     red.setFont(font2);
     
     CellStyle yellow = wb.createCellStyle();
     yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
     yellow.setFillPattern(CellStyle.SOLID_FOREGROUND);
     yellow.setFont(font3);
     
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
	
	 storeIntoList(output2);
	 
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
				 sheet1.autoSizeColumn(iCell);
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
			   	            sheet1.setColumnWidth(iCell,1100*4);
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
						 	if(currentCell.getCellFormula().equals("RIGHT(,10)")){
						 		 rowwrite[i].createCell(iCell).setCellValue("");
						 		continue;
						 	}
		                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getCellFormula());}
					 else if (currentCell.getCellTypeEnum() == CellType.ERROR){
		                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getErrorCellValue());}
			 	}//else part of cell ends
			 
			 }//inner cell for loop ends
		 

			if(i>=6){
				 rowwrite[5].createCell(22).setCellValue("Mobile Check");
				 CellReference cellReference = new CellReference("K"+(i+1));
				 Row rowf = sheet1.getRow(cellReference.getRow());
				 Cell cellf = rowf.getCell(cellReference.getCol());
				 CellValue cellValue = evaluator.evaluate(cellf);
				 
				 try{
				 double val = cellValue.getNumberValue();
				 System.out.println(val);
				 for(int counter = 1;counter<data.size();counter++){
					 if(val==0){System.out.println("HOLA ");continue;}
					 if(val == data.get(counter).getMobile()){
						 rowwrite[i].createCell(22).setCellValue(val);
						 rowwrite[i].createCell(25).setCellValue(data.get(counter).getSource());
						 rowwrite[i].createCell(26).setCellValue(data.get(counter).getID());
						 rowwrite[i].createCell(27).setCellValue(data.get(counter).getcurrentStage());
						 rowwrite[i].createCell(28).setCellValue(data.get(counter).getcurrentStatus());
					 }
				 }
				 
				 }catch(NullPointerException e){}				 
			}

			if(i>=6){
				 rowwrite[5].createCell(23).setCellValue("Email Check");
				 rowwrite[5].createCell(24).setCellValue("Duplicacy Check");
				 rowwrite[5].createCell(25).setCellValue("Source");
				 rowwrite[5].createCell(26).setCellValue("AMS ID");
				 rowwrite[5].createCell(27).setCellValue("Current Stage");
				 rowwrite[5].createCell(28).setCellValue("Current Status");
			     rowwrite[5].createCell(29).setCellValue("Communication");
				 CellReference cellReferences = new CellReference("E"+(i+1));
				 Row rowfs = sheet1.getRow(cellReferences.getRow());
				 Cell cellfs = rowfs.getCell(cellReferences.getCol());
				 CellValue cellValues = evaluator.evaluate(cellfs);
				 
				 try{
				 String vals = cellValues.getStringValue();
				 System.out.println(vals);
				 for(int counter2 = 1;counter2<data.size();counter2++){
					 if(vals.equalsIgnoreCase(data.get(counter2).getEmail())){
						 rowwrite[i].createCell(23).setCellValue(vals);
						 rowwrite[i].createCell(25).setCellValue(data.get(counter2).getSource());
						 rowwrite[i].createCell(26).setCellValue(data.get(counter2).getID());
						 rowwrite[i].createCell(27).setCellValue(data.get(counter2).getcurrentStage());
						 rowwrite[i].createCell(28).setCellValue(data.get(counter2).getcurrentStatus());
					 }
				 }
				 
				 }catch(NullPointerException e){}				 
			}
			
			
		}//row not null ends
		 System.out.println("WorkBook has been created");
		 
//setting the font as bold for the headers
		 
	      if(i>=5){		    
	    	  for(int x =0;x<rowwrite[5].getLastCellNum();x++){
		   Cell co = sheet1.getRow(5).getCell(x);
			co.setCellStyle(style);
	      }}
	      
	      if(i>=6){
	    	 
	    	  Row are = sheet1.getRow(i);
	    	  Cell toCheck1 = are.getCell(22);
	    	  if(toCheck1!=null){
	    		  check = 1;
	    	  }
	    	  Cell toCheck2 = are.getCell(23);
	    	  Row ar = sheet1.getRow(9);
	    	  
	    	  try{	  
	    		  if(toCheck2.getCellTypeEnum()==CellType.STRING){
		    		  Cell color = are.createCell(24);
		    		   	color.setCellValue("Duplicate");
		    		   	color.setCellStyle(red);
		    		   	check = 0;
		    		   	continue;
	    		  }
	    		  else if(toCheck1.getCellTypeEnum()==CellType.NUMERIC&&toCheck2.getCellTypeEnum()==CellType.STRING){
	    		  Cell color = are.createCell(24);
	    		   	color.setCellValue("Duplicate");
	    		   	color.setCellStyle(red);
	    		   	check = 0;
	    		   	continue;
	    	  }
	    	  }catch(NullPointerException s){
	    		  if(check==1){
	    			  Cell color = are.createCell(24);
		    		   	color.setCellValue("To Check");
		    		   	color.setCellStyle(yellow);
		    		   	check = 0;
		    		   	continue;
	    		  }
	    		  
	    		  Cell color = are.createCell(24);
	    		   	color.setCellValue("Unique");
	    		   	color.setCellStyle(green);
	    		   	check = 0;	
	    		   	continue;
	    	  }
	      }
	         

}//outer for loop end
	 for(int i=6;i<sheet1.getLastRowNum();i++){
	     Row ttya = sheet1.getRow(i);
	 
	   try{
	     Cell celler = ttya.getCell(25);
	     String val = celler.getStringCellValue();
	     System.out.println("Value: "+celler.getStringCellValue());
	     if(celler.getStringCellValue().equals("Employee Referral")||celler.getStringCellValue().equals("Repository Activation - P")||celler.getStringCellValue().equals("Repository Activation MR - P")){
	    	  ttya.createCell(25).setCellValue(val);
	     }
	     else{
	   	  ttya.createCell(25).setCellValue("");
	     }
	   }catch(NullPointerException er){}


	   }
	 
	 for(int pounter = 19;pounter<30;pounter++){
		 if(pounter==22){
         	sheet1.setColumnWidth(pounter,1200*4);
         	continue;
         }
            sheet1.setColumnWidth(pounter,1700*4);
	 }
	 
//Writing to the file
	 	FileOutputStream fileOut = new FileOutputStream("AmsDumpOutput"+open+".xlsx");
	    wb.write(fileOut);; 
	    
		fileOut.close();
	    wbread.close();
		System.out.println("Sheet1 of WorkBook has been created");
		System.out.println(data.get(81454).getcurrentStage() + data.get(81454).getPan() + data.get(81454).getID());
		System.out.println(data.get(81463).getcurrentStatus() + data.get(81463).getPan() + data.get(81463).getID());
		File output = new File("AmsDumpOutput.xlsx");
		String path = output.getPath();
		Runtime.getRuntime().exec("explorer.exe /select," + path);
	wb = null;
	myStream = null;

}

public void storeIntoList(String output2){
	File is = new File(output2);
	Workbook workbook = StreamingReader.builder()
						.rowCacheSize(100)
						.bufferSize(1096)
						.open(is);
	int cols = 0;int rows = 0;
	  for (Sheet sheet : workbook){
		  System.out.println(sheet.getSheetName() + " Rows :  "+sheet.getLastRowNum());
	    	for (Row r : sheet) {
    			System.out.println(r.getRowNum());
	    		DataStorer obj = new DataStorer();
	    		for (Cell c : r) {
	    			if(c.getCellTypeEnum()==CellType.NUMERIC){
	    				if(cols==0){
	        				obj.setMobile(c.getNumericCellValue());
	        				cols++;
	        				continue;}
	        				if(cols==2){
	        					obj.setID(c.getNumericCellValue());
	        					cols++;
	        				continue;}
	        			}else{
	    //if the content of the second input sheet is String write that String to the output workbook
	        			if(cols==1){
	        				obj.setEmail(c.getStringCellValue());
	    	    		cols++;}
	        			else if(cols==3){
	        				obj.setPan(c.getStringCellValue());
	        				cols=4;
	        				continue;
	        	    		}
	        			if(cols==4){
	        				obj.setSource(c.getStringCellValue());
	        				cols =5;
	        				continue;
	        			}
	        			if(cols==5){
	        				obj.setcurrentStage(c.getStringCellValue());
	        				cols=6;
	        				continue;
	        			}
	        			if(cols==6){
	        				obj.setcurrentStatus(c.getStringCellValue());
	        				cols++;
	        				continue;
	        			}
	    	    		
	        			}

	    		}
	        		data.add(obj);
	    	      rows++;;//iterating the row to next row
	    	      
	    	      cols=0;//setting the columns again to zero
	    	    				}
	    	}
	  workbook = null;
	  is = null;
	  String email,pan="Nil",source,stage,status;
	  double mobile,ID;
	  for(int counter = 0;counter<data.size();counter++){
		  DataStorer obj = new DataStorer();
		  if(data.get(counter).getcurrentStatus()==null){
			  source = data.get(counter).getPan();
			  stage = data.get(counter).getSource();
			  status = data.get(counter).getcurrentStage();
			  mobile = data.get(counter).getMobile();
			  email = data.get(counter).getEmail();
			  ID = data.get(counter).getID();
			  obj.setMobile(mobile);obj.setcurrentStage(stage);obj.setcurrentStatus(status);
			  obj.setEmail(email);obj.setID(ID);obj.setPan(pan);obj.setSource(source);
			  data.set(counter,obj);
		  }
	  }  
}

public class DataStorer{
	double mobile;
	String email;
	double ID;
	String Pan;
	String Source,Stage,Status;
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
	public void setSource(String source){
		this.Source = source;
	}
	
	public String getSource(){
		return this.Source;
	}
	public void setcurrentStage(String stage){
		this.Stage=stage;
	}
	public String getcurrentStage(){
	return this.Stage;
	}
	public void setcurrentStatus(String status){
		this.Status=status;
	}
	public String getcurrentStatus(){
		return this.Status;
	}
}
//main function to call the constructor

	public static void main(String[] args) throws IOException,InvalidFormatException{
		output1 = "VLookupOutputs.xlsx";	//first input file to be passed
		output2 ="AMS_Dump_Data"+open+"x.xlsx";  //second input file to be passed
		new AMSdump(output1,output2); 		   //calling the constructor
	}
}
