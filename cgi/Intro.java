package cgi;
import java.awt.Color;
import java.awt.event.*;
import java.io.*;
import javax.swing.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

public class Intro{
JLabel file_selected;
JFrame f;
JLabel ta;
JProgressBar jb;    
int i=0,num=0;     
//constructor

Intro() throws IOException {
		f = new JFrame("Beta One");
		JButton b = new JButton("Select File");
		file_selected = new JLabel();
		ta = new JLabel("Select the File for the Master Tracker");
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
		b.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
        				
        	//provide user to select the file
        	 String result = selectfile();
        	 if(result==null){
        	 ta.setText("No File Choosen");
        	 return;}
        	 try {
        		 read_write(result);
				 }catch(IOException e1) {}
        	}
     });
	
}

//file chooser option to select a file for master tracker
public String selectfile(){
//Jfilechooser is used
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

//to read the Master tracker and apply the modifications to generate a new file
public void read_write(String result) throws IOException{
	//to write a new formatted Master Tracker
	Workbook wbwrite = new HSSFWorkbook();
	CreationHelper createHelper = wbwrite.getCreationHelper();
	
	Sheet sheet_write = wbwrite.createSheet("Sheet1");
	Sheet sheet_write2 = wbwrite.createSheet("Sheet2");
	
	FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();
	//to read Master tracker from the file selected by the user
	
	FileInputStream myStream = new FileInputStream(result);
	NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
	HSSFSheet sheet = wb.getSheetAt(0);
	HSSFRow row;
	HSSFCell cell;
	int fCell,lCell;
	int rowStart = sheet.getFirstRowNum();
	int rowEnd =   sheet.getLastRowNum();
	Row rowwrite[] =new Row[rowEnd+1];
	System.out.println(rowStart + "  "+rowEnd);
	String s1 = "",s2="";
    int counter1 = 7;
    int counter2 = 7;	    
	//font style to set font as bold
	CellStyle style = wbwrite.createCellStyle();
	Font font = wbwrite.createFont();font.setFontHeightInPoints((short)11);
	font.setFontName(HSSFFont.FONT_ARIAL);
	font.setBold(true);
	style.setFont(font); 
	//code to iterate over the rows  
	for(int i=rowStart;i<=rowEnd;i++){
	row=sheet.getRow(i);
	if(row==null){
		System.out.println("empty accessed");
		continue;
	}
	if(row!=null){
		rowwrite[i]=sheet_write.createRow((short)i);
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
			 sheet_write.autoSizeColumn(iCell);
		 // testing for types of the cell
			 DataFormatter dataFormatter = new DataFormatter();
			// String cellStringValue = dataFormatter.formatCellValue(row.getCell(iCell));
		   	// rowwrite[i].createCell(iCell+1).setCellValue(cellStringValue);	 
			 if(i>=6 && iCell==5 ||i>=6 && iCell==6){
				 CellStyle dateStyle = wbwrite.createCellStyle();
	    		 dateStyle.setDataFormat(
	    		 createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
	    		 Cell writeDate = rowwrite[i].createCell(iCell+1);
	   	         writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
	   	         writeDate.setCellStyle(dateStyle); 
	   	         sheet_write.setColumnWidth(iCell,1100*4);
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
			 s1=""+cell;
			 s2 += s1 + "\t";}}//cell for loop ends
             s2+= "\n";
         //Validation Index Calculation
         	if(i>=6){
         		rowwrite[i]=sheet_write.getRow((short)i);;
         		rowwrite[i].createCell(0).setCellFormula("CONCATENATE(F"+counter1+",D"+counter2+")");
         		CellReference cellReference = new CellReference("A"+counter1);
         		Row rowF = sheet_write.getRow(cellReference.getRow());
         		Cell cellF = rowF.getCell(cellReference.getCol()); 
         		System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
         		CellValue cellValue = evaluator.evaluate(cellF);
         		System.out.println("  "+cellValue.getStringValue());
         		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
         		counter1+=1;counter2+=1;}
         	if(i==5){
         		rowwrite[i]=sheet_write.getRow((short)i);;
         		rowwrite[i].createCell(0).setCellValue("Validation Index");}}//row not null ends
	
	  String ss = result;
	 System.out.println(result);
      System.out.println("WorkBook has been created");
      }//row ends
	
	  FileOutputStream fileOut = new FileOutputStream(result+"(formatted).xls");
      wbwrite.write(fileOut);
      fileOut.close();
	  ta.setText(s2);
	  wbwrite.close();
	  wb.close();
	  fs.close();
	  finish();
}

public void showProgress(){
	jb=new JProgressBar(0,2000);    
	jb.setBounds(340,40,160,30);         
	jb.setValue(0);    
	jb.setStringPainted(true);    
	f.add(jb);
	iterate();
}

public void iterate(){    
while(i<=2000){    
  jb.setValue(i);    
  i=i+20;    
  try{Thread.sleep(10);}catch(Exception e){}}   
}

public void finish(){
	ta.setText("Finished!");
}



public static void main(String[] args) throws IOException{
	Intro o = new Intro();
}
}
