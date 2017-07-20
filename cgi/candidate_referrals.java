/* Automation of Member Referral Process (June-2017) 
 * Author - Shubham Kumar Singh
 * Email - singh.shubham0812@gmail.com
 * College - Nitte Meenakshi Institute of Technology, Bangalore 
 */
package cgi;

import java.awt.Color;
import java.awt.Toolkit;
import java.awt.event.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javax.swing.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class candidate_referrals {
	//static declarations of the components
	static JLabel file_selected;
	static JFrame f;
	static JTextArea ta;
	
public String selectfile(){
	 //file choosing option
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

//method to format the Candidate Generic File
public void modify(String result) throws IOException{
	int date1x=0,date2x=0,date3=0,date4=0,date5=0;
	int job = 0;String jobA = "";
	int canx = 0;String canxA = "";
	Workbook wbwrite = new XSSFWorkbook();
	CreationHelper createHelper = wbwrite.getCreationHelper();
	Sheet sheet_write = wbwrite.createSheet("new sheet");
	Workbook wbmain = new HSSFWorkbook(); 
	wbmain.getSheet("Sheet 2");
	FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();
    NPOIFSFileSystem fs = new NPOIFSFileSystem(new File(result));
    HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
    HSSFSheet sheet = wb.getSheetAt(0);
    HSSFRow row;
    HSSFCell cell;
    int fCell,lCell;
    int rowStart = sheet.getFirstRowNum(); int rowEnd = sheet.getLastRowNum();
	Row rowwrite[] =new Row[rowEnd+1];
    int counter1 = 7; int counter2 = 7;
    try{ 
    	List<String> heading = new ArrayList<String>();
    	Row extra = sheet.getRow(5);
    	for(int counter=0;counter<extra.getLastCellNum();counter++){
    		Cell extraCell = extra.getCell(counter);
    	    heading.add(extraCell.getStringCellValue());}
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
    	    		job=ca; jobA = Intro.checkAlphabet(job+1);}
	    	if(heading.get(ca).toString().equals("CandidateID")){
    	    		canx=ca;
    	    		canxA = Intro.checkAlphabet(canx+1);}
    	  	}
    	    System.out.println(heading.get(2).toString());
    	    }catch(NullPointerException e){}
    
    
    
	for(int i=rowStart;i<=rowEnd;i++){
		row=sheet.getRow(i);
		if(row==null){
			System.out.println("empty accessed");
			continue;}
		if(row!=null){
			rowwrite[i]=sheet_write.createRow((short)i);
			fCell = row.getFirstCellNum(); 
			lCell = row.getLastCellNum();	
			for (int iCell = fCell; iCell < lCell; iCell++) {
				cell = row.getCell(iCell);
				if(cell==null){
					continue;}
				else{
					//getting reference of current cell
					 Cell currentCell = cell;
					 sheet_write.autoSizeColumn(iCell);
					 if(i>=6 && iCell==date1x||i>=6 && iCell==date2x ||i>=6 && iCell==date3||i>=6 && iCell==date4||i>=6 && iCell==date5 ){
						 try{
							 CellStyle dateStyle = wbwrite.createCellStyle();
							 dateStyle.setDataFormat(
		    		         createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
							 Cell writeDate = rowwrite[i].createCell(iCell+1);
							 writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
							 writeDate.setCellStyle(dateStyle); 
							 sheet_write.setColumnWidth(iCell,1100*4);
							 continue;
							}catch(Exception e){}
					 }
					 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
						 rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());}
					 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue()); }
					 else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
						 rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());}
					 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
	                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());}
			 }
     	}//for ends
	        if(i>=6){
	        	rowwrite[i]=sheet_write.getRow((short)i);;
	        	rowwrite[i].createCell(0).setCellFormula("CONCATENATE("+jobA+counter1+","+canxA+counter2+")");
	        	CellReference cellReference = new CellReference("A"+counter1);
    	 		Row rowF = sheet_write.getRow(cellReference.getRow());
    	 		Cell cellF = rowF.getCell(cellReference.getCol()); 
    	 		CellValue cellValue = evaluator.evaluate(cellF);
    	 		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
	    		counter1+=1;counter2+=1;}
	         	if(i==5){
	         		rowwrite[i]=sheet_write.getRow((short)i);;
	         		rowwrite[i].createCell(0).setCellValue("Validation Index");}
				}		

	            System.out.println("WorkBook has been created");
	}//loop ends
	String path = result.replaceAll(".xls","");
	FileOutputStream fileOut = new FileOutputStream(path+"(Output2).xlsx");
	wbwrite.write(fileOut);
	fileOut.close();    
	wbmain.close();
	wbwrite.close();
	Toolkit.getDefaultToolkit().beep();
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
				return;}
			try {
				o.modify(result);
				} catch (IOException e1) {}
    	}
	});
	}
}
