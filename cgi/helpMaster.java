package cgi;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;  

public class helpMaster implements KeyListener {
	 private static JDialog d; 

	 JLabel heading,l1,l2,l3,l4,l5,l6,l7,l8,l9,l10,l11,l12,l13,l14,l15,l16,l17,l18,l19;
	 helpMaster(){
		  JFrame f= new JFrame();
		  
	        d = new JDialog(f , "Master Referral Validation Automator", true);  
	        d.setLayout(null);  
	       
	        heading = new JLabel("Please Check that the Master Tracker Input File has these fields in any order : ");
	        heading.setBounds(10,20,900,30); 
	        heading.setFont(new Font("Tahoma",Font.PLAIN,20));
	        heading.setForeground(Color.BLUE);
	        l1 = new JLabel(" Title");
	        l1.setBounds(10,80,150,30);
	        l1.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2 = new JLabel(" Candidate Full Name");
	        l2.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2.setBounds(100,80,250,30);
	        l3 = new JLabel(" Candidate ID");
	        l3.setBounds(300,80,150,30);
	        l3.setFont(new Font("Tahoma",Font.BOLD,15));
	        l4 = new JLabel(" Candidate Email");
	        l4.setBounds(450,80,150,30);
	        l4.setFont(new Font("Tahoma",Font.BOLD,15));
	        l5 = new JLabel(" REQ #");
	        l5.setBounds(630,80,150,30);
	        l5.setFont(new Font("Tahoma",Font.BOLD,15));
	        l6 = new JLabel(" Applied Date (WEB)");
	        l6.setBounds(740,80,250,30);
	        l6.setFont(new Font("Tahoma",Font.BOLD,15));
	        l7 = new JLabel(" Applied Date (WEB/MCH)");
	        l7.setBounds(10,120,250,30);
	        l7.setFont(new Font("Tahoma",Font.BOLD,15));
	        l8 = new JLabel(" Business Unit (Hierarchy)");
	        l8.setBounds(240,120,250,30);
	        l8.setFont(new Font("Tahoma",Font.BOLD,15));
	        l9 = new JLabel(" Business Unit (Req More)");
	        l9.setBounds(470,120,250,30);
	        l9.setFont(new Font("Tahoma",Font.BOLD,15));
	        l10 = new JLabel(" Candidate Phone Number");
	        l10.setBounds(700,120,250,30);
	        l10.setFont(new Font("Tahoma",Font.BOLD,15));
	        l11 = new JLabel(" Candidate Source");
	        l11.setFont(new Font("Tahoma",Font.BOLD,15));
	        l11.setBounds(10,160,250,30);
	        l12 = new JLabel(" Candidate Skills");
	        l12.setFont(new Font("Tahoma",Font.BOLD,15));
	        l12.setBounds(190,160,250,30);
	        l13 = new JLabel(" Cell Phone");
	        l13.setFont(new Font("Tahoma",Font.BOLD,15));
	        l13.setBounds(370,160,250,30);
	        l14 = new JLabel(" Cell Telephone");
	        l14.setFont(new Font("Tahoma",Font.BOLD,15));
	        l14.setBounds(500,160,250,30);
	        l15 = new JLabel(" Current Salary Rate");
	        l15.setFont(new Font("Tahoma",Font.BOLD,15));
	        l15.setBounds(670,160,250,30);
	        l16 = new JLabel(" Desired Salary");
	        l16.setFont(new Font("Tahoma",Font.BOLD,15));
	        l16.setBounds(10,200,250,30);
	        l17 = new JLabel(" SBU");
	        l17.setFont(new Font("Tahoma",Font.BOLD,15));
	        l17.setBounds(180,200,250,30);
	        l18 = new JLabel(" Referred By Email");
	        l18.setFont(new Font("Tahoma",Font.BOLD,15));
	        l18.setBounds(270,200,250,30);
	        l19 = new JLabel(" Referred By Name");
	        l19.setFont(new Font("Tahoma",Font.BOLD,15));
	        l19.setBounds(470,200,250,30);
	        JLabel txt = new JLabel("Press any key to Exit...");
	        txt.setBounds(370,250,250,30);
	        txt.setForeground(Color.RED);
	        txt.setFont(new Font("Tahoma",Font.PLAIN,14));
	        
	        
	        
	        
	        d.addKeyListener(this);  
	        d.add(heading);d.add(txt);
	        d.add(l1);d.add(l2);d.add(l3);d.add(l4);d.add(l5);d.add(l6);d.add(l7);d.add(l8);d.add(l9);d.add(l10);
	        d.add(l11);d.add(l12);d.add(l13);d.add(l14);d.add(l15);d.add(l16);d.add(l17);d.add(l18);d.add(l19);
	        d.setSize(940,340);
	        d.setLocation(250,30);
	      
	        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        d.setVisible(true);  
	       
	        
}
	 
	 
public static int checkFormat(String input2){
	 try{int count=0;
	 java.util.List<String> fields = new ArrayList<String>();
		 if(input2.contains(".xls")||input2.contains(".xlsx")){
	     FileInputStream inputStream = new FileInputStream(new File(input2));
	     Workbook workbook = WorkbookFactory.create(inputStream);
	     Sheet sheet = workbook.getSheet("Sheet1");
	     int rowEnd = sheet.getLastRowNum() ;
	         Row ttya = sheet.getRow(5);
	         int colCount = ttya.getLastCellNum();
	         for(int i =0;i<colCount;i++){
	        	 Cell cell = ttya.getCell(i);
	        	 fields.add(cell.getStringCellValue());
	         }
	         	 for(int x = 0;x<fields.size();x++){
	        	 if(fields.get(x).equals("Title")){ count++;}if(fields.get(x).equals("Candidate Full Name")){ count++;}
	        	 if(fields.get(x).equals("Candidate ID")){ count++;}if(fields.get(x).equals("Candidate Email")){ count++;}
	        	 if(fields.get(x).equals("REQ #")){ count++;}if(fields.get(x).equals("Applied Date (WEB)")){ count++;}
	        	 if(fields.get(x).equals("Applied Date (WEB/MCH)")){ count++;}if(fields.get(x).equals("Business Unit (Hierarchy)")){ count++;}
	        	 if(fields.get(x).equals("Business Unit (Req More)")){ count++;}if(fields.get(x).equals("Candidate Phone Number")){ count++;}
	        	 if(fields.get(x).equals("Candidate Source")){ count++;}if(fields.get(x).equals("Candidate Skills")){ count++;}
	        	 if(fields.get(x).equals("Cell Phone")){ count++;}if(fields.get(x).equals("Cell telephone")){ count++;}
	        	 if(fields.get(x).equals("Current Salary Rate")){ count++;}if(fields.get(x).equals("Desired Salary")){ count++;}
	        	 if(fields.get(x).equals("SBU")){ count++;}if(fields.get(x).equals("Referred By Email")){ count++;}
	        	 if(fields.get(x).equals("Referred By")){ count++;}
	         	 }
	        	 if(count==19){
	        		return count;
	        	 }
	        	 return count;
		 }
		 else{
			 throw new IOException();
		 }
	 }catch (IOException | EncryptedDocumentException
	         | InvalidFormatException ex) {
     JOptionPane.showMessageDialog(null, "Error : Invalid File Selected");
     return -1;
	 }
}
	 
	 public static void main(String...args){
		 new helpMaster();
		 String x ="Njoyn Master Tracker-16-18.xls";
		 checkFormat(x);
	 }

	@Override
	public void keyPressed(KeyEvent e) {
		d.dispose();
	}

	@Override
	public void keyReleased(KeyEvent arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void keyTyped(KeyEvent arg0) {
		}
		
}
