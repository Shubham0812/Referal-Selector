package cgi;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextArea;

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
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class candidate_referrals {
	static JLabel file_selected;
	static JFrame f;
	static JTextArea ta;
	//constructor
	
	candidate_referrals() throws IOException{

	}
public String selectfile(){
	 //file chooser
    JFileChooser fileChooser = new JFileChooser();
fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
int result = fileChooser.showOpenDialog(f);
if (result == JFileChooser.APPROVE_OPTION) {
    File selectedFile = fileChooser.getSelectedFile();
    String filePath = selectedFile.getPath();
    file_selected.setText("Selected file: " + filePath);
	    return filePath;
	}
	return null;
}

public void modify(String result) throws IOException{
	
	int date1=0,date2=0,date3=0,date4=0,date5=0;
	int job = 0;String jobA = "";
	int can = 0;String canA = "";
	Workbook wbwrite = new XSSFWorkbook();
	CreationHelper createHelper = wbwrite.getCreationHelper();
	Sheet sheet_write = wbwrite.createSheet("new sheet");
	
	Workbook wbmain = new HSSFWorkbook(); 
	wbmain.getSheet("Sheet 2");
	
	FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();
	FileInputStream myStream = new FileInputStream(result);
    NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
    HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
    HSSFSheet sheet = wb.getSheetAt(0);

    HSSFRow row;
    HSSFCell cell;
    int fCell,lCell;
    int rowStart = sheet.getFirstRowNum();
    int rowEnd = sheet.getLastRowNum();
	Row rowwrite[] =new Row[rowEnd+1];
    System.out.println(rowStart + "  "+rowEnd);
    int counter1 = 7;
    int counter2 = 7;
    
    try{ 
    	 List<String> headings = new ArrayList<String>();
    	 Row extra = sheet.getRow(5);
    	    for(int counter=0;counter<extra.getLastCellNum();counter++){
    	        Cell extraCell = extra.getCell(counter);
    	        headings.add(extraCell.getStringCellValue());
    	        }
    	    for(int ca = 0;ca<headings.size();ca++){
    	    	
    	    	if(headings.get(ca).toString().equals("Application Date")){
    	    		date1=ca;
    	    		System.out.println(ca);
    	    	}
    	    	if(headings.get(ca).toString().equals("Date Survey Taken")){
    	    		date2=ca;
    	    		System.out.println(ca);
    	    	}
    	    	if(headings.get(ca).toString().equals("Date Survey Invite Sent")){
    	    		date3=ca;
    	    		System.out.println(ca);
    	    	}
    	    	if(headings.get(ca).toString().equals("Candidate Enter Date")){
    	    		date4=ca;
    	    		System.out.println(ca);
    	    	}
    	    	if(headings.get(ca).toString().equals("Last Activity Date")){
    	    		date5=ca;
    	    		System.out.println(ca);
    	    	}
    	    	if(headings.get(ca).toString().equals("Job ID")){
    	    		job=ca;
    	    		jobA = Intro.checkAlphabet(job+1);
    	    		System.out.println(job+jobA);
    	    		
    	    	}
    	    	if(headings.get(ca).toString().equals("CandidateID")){
    	    		can=ca;
    	    		canA = Intro.checkAlphabet(can+1);
    	    		System.out.println(can+canA);
    	    		
    	    	}
    	    		
    	    	
    	    }
    	    System.out.println(headings.get(2).toString());
    	    }catch(NullPointerException e){}
    
    
    
	for(int i=rowStart;i<=rowEnd;i++){
	row=sheet.getRow(i);
	if(row==null){
		System.out.println("empty accessed");
		continue;
		}
	if(row!=null){
		rowwrite[i]=sheet_write.createRow((short)i);
		 fCell = row.getFirstCellNum(); 
         lCell = row.getLastCellNum();	
         for (int iCell = fCell; iCell < lCell; iCell++) {
			 cell = row.getCell(iCell);
			 if(cell==null){
				 continue;
			 }
			 else{
					//getting reference of current cell
					 Cell currentCell = cell;
					 sheet_write.autoSizeColumn(iCell);
					 
					 if(i>=6 && iCell==date1||i>=6 && iCell==date2 ||i>=6 && iCell==date3||i>=6 && iCell==date4||i>=6 && iCell==date5 ){
						 try{
		    			 CellStyle dateStyle = wbwrite.createCellStyle();
		    		       dateStyle.setDataFormat(
		    		           createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
		    		       Cell writeDate = rowwrite[i].createCell(iCell+1);
		   	            writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
		   	            writeDate.setCellStyle(dateStyle); 
		   	         sheet_write.setColumnWidth(iCell,1100*4);
		        continue;}catch(Exception e){}
					 }
					 
					 
					 
					 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
	                     System.out.print(currentCell.getNumericCellValue() + "--");
	                     
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
					 }
					 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
	                     System.out.print(currentCell.getStringCellValue() + "--");
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());    
					 }
					 else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
	                     System.out.print(currentCell.getStringCellValue() + "--");
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());    
					 }
					 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
	                    System.out.print(currentCell.getStringCellValue() + "--");
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());    
					 	}
			 }
     	}//for ends
	         	if(i>=6){
	           	 rowwrite[i]=sheet_write.getRow((short)i);;
	    		 rowwrite[i].createCell(0).setCellFormula("CONCATENATE("+jobA+counter1+","+canA+counter2+")");
	    		 CellReference cellReference = new CellReference("A"+counter1);
	    	 		Row rowF = sheet_write.getRow(cellReference.getRow());
	    	 		Cell cellF = rowF.getCell(cellReference.getCol()); 
	    	 		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
	    	 		CellValue cellValue = evaluator.evaluate(cellF);
	    	 		System.out.println("  "+cellValue.getStringValue());
	    	 		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
	    		 counter1+=1;counter2+=1;
	            }
	         	if(i==5){
	         		rowwrite[i]=sheet_write.getRow((short)i);;
	         		 rowwrite[i].createCell(0).setCellValue("Validation Index");
	        	}
				}		

	            System.out.println("WorkBook has been created");
	}//row ends
	String path = result.replaceAll(".xls","");
	  FileOutputStream fileOut = new FileOutputStream(path+"(Output2).xlsx");
      wbwrite.write(fileOut);
      fileOut.close();    
      wbmain.close();
      wbwrite.close();
      fs.close();
      wb.close();
 }
public static void main(String[] args) throws IOException{
	f = new JFrame("Beta One");
	JButton b = new JButton("Select File");
	file_selected = new JLabel();
	 ta = new JTextArea("Select the File for the Candidate Referral");
	b.setBounds(10, 30, 150, 30);
	file_selected.setBounds(550, 130, 500, 30);
	ta.setBounds(200,20,600,600); 
	ta.setBackground(new Color(255,255,255));
	f.add(b);f.add(file_selected);f.add(ta);
	f.setSize(1000,500);
	f.setLayout(null);
	f.getContentPane().setBackground(new Color(255,255,255));
	//f.setLocation(400,100);
	f.setVisible(true);
	f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	candidate_referrals o = new candidate_referrals();
	b.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
        			//provide user to select the file
        			String result = o.selectfile();
        			if(result==null){
        				ta.setText("No File Choosen");
        				return;
        			}
        			try {
						o.modify(result);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
        	}
	});
	}
	
}
