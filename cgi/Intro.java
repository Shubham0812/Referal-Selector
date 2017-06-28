package cgi;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JProgressBar;
import javax.swing.SwingConstants;

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
import java.awt.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Intro{
JFrame f;
JLabel ta,label1,label2,label3,error;
JButton b,ba,b1,b2,b3,submit,Vb1,Vb2;
JProgressBar jb;    
String result,result2;
String outputFile1,outputFile2;
int i=0,num=0,count=0;    
int code;
boolean both_set = false;
//constructor

Intro() throws IOException {
		JFrame.setDefaultLookAndFeelDecorated(true);	
		f = new JFrame("Master Tracker Generator");
		f.setLayout(null);
		b = new JButton("Select Master Tracker File");
		b.setBounds(460, 130, 350, 30);
		ba = new JButton("Select Candidate (Generic) File");
		ba.setBounds(460, 200, 350, 30);
		ba.setVisible(false);
		b.setVisible(false);
		b1 = new JButton("Module 1: ");
		b1.setBounds(10, 130, 150,30);
		b2 = new JButton("Module 2: ");
		b2.setBounds(10, 200, 150,30);
		b3 = new JButton("Module 3: ");
		b3.setBounds(10, 270, 150, 30);
		submit = new JButton("Submit Selected File");
		submit.setBounds(370,320,200,30);
		submit.setVisible(false);
		Vb1= new JButton("Select Master Tracker File");
		Vb1.setBounds(460,270,210,30);
		Vb1.setVisible(false);
		Vb2= new JButton("Select Candiate Referral File");
		Vb2.setBounds(680,270,210,30);
		Vb2.setVisible(false);
		label1 = new JLabel("Format Master Tracker");
		label1.setBounds(190, 130, 500, 30);
		label1.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label2 = new JLabel("Format Candidate Referral");
		label2.setBounds(190, 200, 500, 30);
		label2.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label3 = new JLabel("Perform VLookup");
		label3.setBounds(190, 270, 500, 30);
		label3.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		error = new JLabel("");
		error.setBounds(400, 430, 350, 30);
		ta = new JLabel("Member Referral Validation Automator",SwingConstants.CENTER);
		ta.setBounds(200,0,600,100); 
		ta.setFont(new Font("Courier New", Font.BOLD, 22));
		f.add(b);f.add(label1);f.add(ta);f.add(b1);f.add(b2);f.add(b3);f.add(ba);f.add(label2);f.add(label3);f.add(error);f.add(Vb1);f.add(Vb2);
		f.add(submit);
		f.setSize(1000,500);
		f.getContentPane().setBackground(new Color(255,255,255));
		f.setLocation(240,20);
		f.setVisible(true);
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		b1.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        b.setVisible(true);
	        error.setText("");
	        b2.setEnabled(false);
	        b3.setEnabled(false);}			
	     });
		b2.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        error.setText("");
	        ba.setVisible(true);
	        b1.setEnabled(false);
	        b3.setEnabled(false);}			
	     });
		
		b3.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        try {
		        b1.setEnabled(false);
		        b2.setEnabled(false);
		        Vb1.setVisible(true);
		        Vb2.setVisible(true);
				//new SheetCopy();
				//finish();
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}}			
	     });
		
		Vb1.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        	  outputFile1 = selectfile();
	        	 if(outputFile1==null){
	        	 error.setText("Master Tracker File Not Choosed");
	        	 count+=1;
	        	 return;
	        	 }
	        	 	error.setText("Master Tracker File Selected");
	        }
	     });
		Vb2.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        	  outputFile2 = selectfile();
	        	 if(outputFile2==null){
	        	 error.setText("Candidate Referral File Not Choosed");
	        	 count+=1;
	        	 return;
	        	 }
	        	 error.setText("Candidate Referral File Selected");
	        	 if(outputFile1!=null&&outputFile2!=null){
	        		   code=3;
	        		 submit.setVisible(true);
	        	        	 }
					 }
	     });
		b.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
        				error.setText("");
        	//provide user to select the file
        	  result = selectfile();
        	 if(result==null){
        	 error.setText("No File Choosen");
        	 b.setVisible(false);
        	 b2.setEnabled(true);
        	 b3.setEnabled(true);
        	 return;
        	 }
        		 code=1;
        		 submit.setVisible(true);
        		 //read_write(result);
				 }
     });
		
		ba.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				
	        	//provide user to select the file
	        	 result2 = selectfile();
	        	 if(result2==null){
	        	 error.setText("No File Choosen");
	        	 ba.setVisible(false);
	        	 b1.setEnabled(true);
	        	 b3.setEnabled(true);
	        	 return;
	        	 }
	        		 code=2;
	        		 submit.setVisible(true);

					 }
	     });
		
		submit.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	
	        	if(code==1){
	        	try {
					read_write(result);
	        	} catch (IOException e1) {}
	        			   }
	        	
	        	if(code==2){
	        		try{
	        		 candidate_referrals obj = new candidate_referrals();
	        		 obj.modify(result2);
	        		 finish();
	        		 submit.setVisible(false);
	        		 b1.setEnabled(true);
	        		 b3.setEnabled(true);
	        		 ba.setVisible(false);
	        		}catch(Exception code2){
	        			error.setText("Invalid File Selected");
	        			ba.setVisible(false);
	               	 	b1.setEnabled(true);
	               	 	b2.setEnabled(true);
	               	 	b3.setEnabled(true);
	        		}
	        	}
	        	if(code==3){
	        		try{
	        		new SheetCopy(outputFile1,outputFile2);
	        		finish();
		        	 Vb1.setVisible(false);
		        	 Vb2.setVisible(false);
		        	 b1.setEnabled(true);
		        	 b2.setEnabled(true);
		        	 b3.setEnabled(true);  
		        	 submit.setVisible(false);
	        		
	        		}catch(Exception Vlook){
	        			
	        		}
	        	}
	      }
	     });
		

		
		
	
}

public void checkButtonPress(String result,int code){
	
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
	wbwrite.createSheet("Sheet2");
	
	FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();
	//to read Master tracker from the file selected by the user
	try{
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
    int counter1 = 7;
    int counter2 = 7;	    
	//font style to set font as bold
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
    					 }}
    				   	 }
		 				}
		 //if the cell has value determine the type of value.
		 else{
		 //getting reference of current cell
			 Cell currentCell = cell;
			 sheet_write.autoSizeColumn(iCell);
			 if(i>=6 && iCell==5 ||i>=6 && iCell==6){
				 try{
				 CellStyle dateStyle = wbwrite.createCellStyle();
	    		 dateStyle.setDataFormat(
	    		 createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
	    		 Cell writeDate = rowwrite[i].createCell(iCell+1);
	   	         writeDate.setCellValue(row.getCell(iCell).getDateCellValue());
	   	         writeDate.setCellStyle(dateStyle); 
	   	         sheet_write.setColumnWidth(iCell,1100*4);
	   	         continue;
				 }catch(Exception ex){}
			 }
			 
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
    					 try{
        				 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
        				 }catch(Exception af){}
    					 }
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
	
	 System.out.println(result);
      System.out.println("WorkBook has been created");
      }//row ends
	  String path = result.replaceAll(".xls","");
	  FileOutputStream fileOut = new FileOutputStream(path+"(Output1).xls");
      wbwrite.write(fileOut);
      fileOut.close();
      
      File look = new File(path+"(Output1).xls");
	  String output = look.getPath();
      Runtime.getRuntime().exec("explorer.exe /select," + output);

 	 b.setVisible(false);
 	 submit.setVisible(false);
 	 b2.setEnabled(true);
 	 b3.setEnabled(true);
	  wbwrite.close();
	  wb.close();
	  fs.close();
	  finish();
	}catch(Exception e)
	{
		error.setText("Invalid File Selected");
		b.setVisible(false);
   	 	b2.setEnabled(true);
   	 	b3.setEnabled(true);
		return;
		}
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
	error.setText("Finished!");
}



public static void main(String[] args) throws IOException{
new Intro();
}
}
