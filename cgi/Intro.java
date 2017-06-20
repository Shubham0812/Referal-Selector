package cgi;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;

import javax.swing.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
public class Intro{
	
Intro(){
	JFrame f = new JFrame("Beta One");
	JButton b = new JButton("Select File");
	JLabel file_selected = new JLabel();
	JTextArea ta = new JTextArea("Select the File for the Master Tracker");
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
	
	b.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
        	JFileChooser fileChooser = new JFileChooser();
        	fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
        	int result = fileChooser.showOpenDialog(f);
        	if (result == JFileChooser.APPROVE_OPTION) {
        	    File selectedFile = fileChooser.getSelectedFile();
        	    String filePath = selectedFile.getPath();
        	    file_selected.setText("Selected file: " + filePath);
        	    try{ 
        	    	FileInputStream myStream = new FileInputStream(filePath);
        	   	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
        	   	 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
        	   	 HSSFSheet sheet = wb.getSheetAt(0);
        		 HSSFRow row;
        		 HSSFCell cell;
        		 int rows; // No of rows
        		 rows = sheet.getPhysicalNumberOfRows();
        		 System.out.println(rows);
        		 int cols = 0; // No of columns
        		 int tmp = 0;String s1 = "",s2="";
        		 for(int i=0;i<=rows;i++){
        			 row=sheet.getRow(i);
        			 if(row!=null){
        				 tmp = sheet.getRow(i).getPhysicalNumberOfCells();
        				 System.out.println(tmp);
        				 if(tmp>cols)
        					 cols = tmp;
        				 for(int r = 0;r<=row.getLastCellNum();r++){
        					 cell = row.getCell(r);
        					 if(cell==null){
        						 continue;
        					 }
        					 s1=""+cell.getAddress();
        					 s2 += s1 + "\n";
        				 }
        			 }
        		 }
        		 ta.setText(s2);
        	    }catch (Exception ex) {ex.printStackTrace();  }               
        	}
        }
     });
	
}
public static void main(String[] args){
	new Intro();
}
}
