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
import javax.swing.JTextArea;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
public class Intro{
	
JLabel file_selected;
JFrame f;
JTextArea ta;
//constructor
Intro() throws IOException{
	f = new JFrame("Beta One");
	JButton b = new JButton("Select File");
	file_selected = new JLabel();
	 ta = new JTextArea("Select the File for the Master Tracker");
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
        				return;
        			}
        			try {
						read_write(result);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
        	}
     });
	
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


public void read_write(String result) throws IOException{
	
	
		//to write
		Workbook wbwrite = new HSSFWorkbook();
		CreationHelper createHelper = wbwrite.getCreationHelper();
		Sheet sheet_write = wbwrite.createSheet("new sheet");
	
		//to read
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
	    int cols = 0; // No of columns
	    String s1 = "",s2="";
	    
	    //font set
	       CellStyle style = wbwrite.createCellStyle();
	       Font font = wbwrite.createFont();
	       font.setFontHeightInPoints((short)11);
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
		rowwrite[i]=sheet_write.createRow((short)i);;

		//first and last cell for the row
		 fCell = row.getFirstCellNum(); 
         lCell = row.getLastCellNum();	
         for (int iCell = fCell; iCell < lCell; iCell++) {
			 cell = row.getCell(iCell);
			 if(cell==null){
				 if(iCell==9){
					 Cell currentCells = row.getCell(iCell+3);
					 
					  if(currentCells==null){
					  Cell currentCeller = row.getCell(iCell+4);
		    		  if(currentCeller.getCellTypeEnum() == CellType.NUMERIC){
		    		  double value = currentCeller.getNumericCellValue();
		    		  rowwrite[i].createCell(9).setCellFormula("RIGHT("+value+",10)");
		    		  continue;
							 }
		    		  else if(currentCeller.getCellTypeEnum() == CellType.STRING){
	    					 String value = currentCeller.getStringCellValue();
	    					 String newValue = value.replaceAll("-","");
	    					 rowwrite[i].createCell(9).setCellFormula("RIGHT("+newValue+",10)");
	    					 continue;
	    				 }
		    		  continue;
					 	}
    				 if(currentCells.getCellTypeEnum() == CellType.NUMERIC){
    					 double value = currentCells.getNumericCellValue();
    					 rowwrite[i].createCell(9).setCellFormula("RIGHT("+value+",10)");
    					 continue;
					 }
    				 else if(currentCells.getCellTypeEnum() == CellType.STRING){
    					 String value = currentCells.getStringCellValue();
    					 String newValue = value.replaceAll("-","");
    					 rowwrite[i].createCell(9).setCellFormula("RIGHT("+newValue+",10)");
    				 }
				 continue;
				 		}
				 
			 }//if the cell has value determine the type of value.
			 else{
				//getting reference of current cell
				 Cell currentCell = cell;
				 sheet_write.autoSizeColumn(iCell);
				 //testing for types of the cell
				 
				 DataFormatter dataFormatter = new DataFormatter();
				 String cellStringValue = dataFormatter.formatCellValue(row.getCell(iCell));
				 rowwrite[i].createCell(iCell).setCellValue(cellStringValue);
				 
				 if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                     System.out.print(currentCell.getNumericCellValue() + "--");
                     
    				 if(i>=6&& iCell ==9){
    					 double value = currentCell.getNumericCellValue();
    					 rowwrite[i].createCell(9).setCellFormula("RIGHT("+value+",10)");
    					 continue;
    				 }
                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getNumericCellValue());    
				 }
				 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
                     System.out.print(currentCell.getStringCellValue() + "--");
    				 if(i>=6&& iCell ==9){
    					 String value = currentCell.getStringCellValue();
    					 String newValue = value.replaceAll("-","");
    					 rowwrite[i].createCell(9).setCellFormula("RIGHT("+newValue+",10)");
    					 continue;
    				 }
                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getStringCellValue());    
				 }
				 else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
                     System.out.print(currentCell.getStringCellValue() + "--");
                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getCellFormula());    
				 }
				 else if (currentCell.getCellTypeEnum() == CellType.ERROR) {
                    System.out.print(currentCell.getStringCellValue() + "--");
                     rowwrite[i].createCell(iCell).setCellValue(currentCell.getErrorCellValue());    
				 	}
				 
			 s1=""+cell;
			 s2 += s1 + "\t";
			 	}
	}
         s2+= "\n";
		 	}
	  FileOutputStream fileOut = new FileOutputStream("(formatted).xls");
      wbwrite.write(fileOut);
      fileOut.close();
      System.out.println("WorkBook has been created");
	 }
	//row ends
 ta.setText(s2);
}
public static void main(String[] args) throws IOException{
	new Intro();
}
}
