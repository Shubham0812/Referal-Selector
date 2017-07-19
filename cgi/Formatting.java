package cgi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;

import cgi.AMSdump.DataStorer;

public class Formatting {
	  XSSFWorkbook wb = new XSSFWorkbook();
	 Sheet sheet1 = wb.createSheet("Sheet1");
	 java.util.List<DataStorer> data = new ArrayList<DataStorer>();
	 FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	 
Formatting(String input1) throws IOException, InvalidFormatException{	
	
	 CellStyle style = wb.createCellStyle();
	 Font font = wb.createFont();
     font.setFontHeightInPoints((short)11);
     font.setFontName(HSSFFont.FONT_ARIAL);
     font.setBold(true);
     style.setFont(font);
     
     try{
     FileInputStream inputStream = new FileInputStream(new File(input1));
     Workbook workbook = WorkbookFactory.create(inputStream);
     Sheet sheet = workbook.getSheet("Sheet1");
     int rowCount = sheet.getLastRowNum();
     int rowStart = sheet.getFirstRowNum() ;
     int rowEnd = sheet.getLastRowNum() ;
     System.out.println(rowCount+ "First Row : "+rowStart+"\t Last Row : "+rowEnd);
     
for(int i=rowStart;i<=sheet.getLastRowNum()+1;i++){
     Row ttya = sheet.getRow(i);
   try{
     Cell celler = ttya.getCell(24);
     String val = celler.getStringCellValue();
     System.out.println("Value: "+val);
     if(val.equals("Unique")){
    	 ttya.createCell(29).setCellValue("Comm1");
    	 continue;
     }
     if(val.equals("Duplicate")||val.equals("To Check")){
 
    	 Cell cell = ttya.getCell(25);
    	 String vale = cell.getStringCellValue();
    	 if(vale.equals("")){
    	     		 ttya.createCell(29).setCellValue("Comm2a/2b");
    	 }
     }
   }catch(NullPointerException er){}

try{
	     Cell celler = ttya.getCell(25);
	     String val = celler.getStringCellValue();
	     System.out.println("Value: "+val);
	     if(val.equals("Employee Referral")||val.equals("Repository Activation MR - P")||val.equals("Repository Activation - P")){
	    	 Cell cell = ttya.getCell(27);
	    	 Cell cella = ttya.getCell(28);
	    	 String va = cella.getStringCellValue();
	    	 if(va.contains("Drop")){
	    		 ttya.createCell(29).setCellValue("Comm4a/4b");
	    		 continue;
	    	 }else{
	    	 String data27 = cell.getStringCellValue();
	    	 if(data27.equals("Recruiter CV Screening")||data27.equals("Hiring Manager CV Screening")){
	    	     		 ttya.createCell(29).setCellValue("Comm4a/4b");
	    	 }
	    	 else if(data27.equals("Offer")){
	    			 ttya.createCell(29).setCellValue("Comm3");
	    			 continue;
	    		 }
	    	 else if(data27.equals("First Interview")||data27.equals("Second Interview")){
	    		 if(va.equals("Scheduled")||va.equals("OnHold")||va.equals("Shortlisted")||va.contains("Pending")){
	    			 ttya.createCell(29).setCellValue("Comm3");
	    			 continue;
	    		 }
	    		 else if(va.equals("Candidate No Show")||va.equals("Interview Cancelled")){
	    			 ttya.createCell(29).setCellValue("Comm4a/4b");
	    			 continue;
	    		 }
	    	 }
	    	 else if(data27.equals("Final Interview/Onsite/Client")||data27.equals("Managerial Interview")||data27.equals("HR Interview")){
	    		 Cell cellaa = ttya.getCell(28);
		    	 String va2 = cellaa.getStringCellValue();
		    	 if(va2.equals("Pending")){
		    		 ttya.createCell(29).setCellValue("Comm4a/4b");
	    			 continue;
		    	 }
	    		 ttya.createCell(29).setCellValue("Comm3");
	    		 continue;
	    	 }
	     }
	     }
	   }catch(NullPointerException er){}   
   
   

}//row counter ends

     inputStream.close();
     FileOutputStream outputStream = new FileOutputStream("Ams_Comm_Mails.xlsx");
     workbook.write(outputStream);
     workbook.close();
     outputStream.close();
      
 } catch (IOException | EncryptedDocumentException
         | InvalidFormatException ex) {
     ex.printStackTrace();
 }
}

public static void main(String[] args) throws IOException, InvalidFormatException{
	String input1 = "AmsDumpOutput.xlsx";
	new Formatting(input1);
}
}
