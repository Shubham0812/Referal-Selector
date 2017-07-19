package cgi;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.SwingConstants;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;  

public class helpMaster implements KeyListener {
	 private static JDialog d; 
	 int extra = 0;
	 java.util.List<String> fields = new ArrayList<String>();
	 JLabel heading,l1,l2,l3,l4,l5,l6,l7,l8,l9,l10,l11,l12,l13,l14,l15,l16,l17,l18,l19;
	Boolean j1=false,j2=false,j3=false,j4=false,j5=false,j6=false,j7=false,j8=false,j9=false,j10=false,j11=false,j12=false,j13=false,j14=false,j15=false,j16=false,j17=false,j18=false,j19=false;
	JButton b1,b2;
	public void	justHelp(){
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
	
public void helpMasters(String input){
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
	        l2.setBounds(130,80,250,30);
	        l3 = new JLabel(" Candidate ID");
	        l3.setBounds(320,80,150,30);
	        l3.setFont(new Font("Tahoma",Font.BOLD,15));
	        l4 = new JLabel(" Candidate Email");
	        l4.setBounds(470,80,150,30);
	        l4.setFont(new Font("Tahoma",Font.BOLD,15));
	        l5 = new JLabel(" REQ #");
	        l5.setBounds(640,80,150,30);
	        l5.setFont(new Font("Tahoma",Font.BOLD,15));
	        l6 = new JLabel(" Applied Date (WEB)");
	        l6.setBounds(770,80,250,30);
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
	        b1 = new JButton("Show Error with Fields");
	        b1.setBounds(330,270,230,30);
	        b1.setBackground(new Color(255,82,82));
	        b1.setForeground(new Color(255,255,255));
	        l19.setBounds(470,200,250,30);
	        JLabel txt = new JLabel("Press any key to Exit...");
	        txt.setBounds(375,300,250,30);
	        txt.setForeground(Color.RED);
	        txt.setFont(new Font("Tahoma",Font.PLAIN,14));
		       checkFormat(input);
	        b1.addActionListener(new ActionListener() {
		        public void actionPerformed(ActionEvent e) {
			        checking();}			
			     });
	        
	       
	        d.add(heading);d.add(txt);d.add(b1);
	        d.add(l1);d.add(l2);d.add(l3);d.add(l4);d.add(l5);d.add(l6);d.add(l7);d.add(l8);d.add(l9);d.add(l10);
	        d.add(l11);d.add(l12);d.add(l13);d.add(l14);d.add(l15);d.add(l16);d.add(l17);d.add(l18);d.add(l19);
	        d.setSize(930,380);
	        d.setResizable(false);
	        d.setLocation(250,30);
	      
	        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        d.setVisible(true);  

}
	 
	 
public int checkFormat(String input2){
	 try{int count=0;
	
		 if(input2.contains(".xls")||input2.contains(".xlsx")){
	     FileInputStream inputStream = new FileInputStream(new File(input2));
	     Workbook workbook = WorkbookFactory.create(inputStream);
	     Sheet sheet = workbook.getSheetAt(0);
	     int rowEnd = sheet.getLastRowNum() ;
	         Row ttya = sheet.getRow(5);
	         int colCount = ttya.getLastCellNum();
	         for(int i =0;i<colCount;i++){
	        	 Cell cell = ttya.getCell(i);
	        	 if(cell==null){
	        		 continue;
	        	 }
	        	 fields.add(cell.getStringCellValue());
	         }
	         	 extra = (fields.size()-19);
	         	 for(int x = 0;x<fields.size();x++){
	        	 if(fields.get(x).equals("Title")){j1=true;count++;}if(fields.get(x).equals("Candidate Full Name")){ j2=true;count++;}
	        	 if(fields.get(x).equals("Candidate ID")){ j3=true;count++;}if(fields.get(x).equals("Candidate Email")){ j4=true;count++;}
	        	 if(fields.get(x).equals("REQ #")){ j5=true;count++;}if(fields.get(x).equals("Applied Date (WEB)")){ j6=true;count++;}
	        	 if(fields.get(x).equals("Applied Date (WEB/MCH)")){ j7=true;count++;}if(fields.get(x).equals("Business Unit (Hierarchy)")){ j8=true;count++;}
	        	 if(fields.get(x).equals("Business Unit (Req More)")){ j9=true;count++;}if(fields.get(x).equals("Candidate Phone Number")){ j10=true;count++;}
	        	 if(fields.get(x).equals("Candidate Source")){ j11=true;count++;}if(fields.get(x).equals("Candidate Skills")){ j12=true;count++;}
	        	 if(fields.get(x).equals("Cell Phone")){ j13=true;count++;}if(fields.get(x).equals("Cell telephone")){ j14=true;count++;}
	        	 if(fields.get(x).equals("Current Salary Rate")){ j15=true;count++;}if(fields.get(x).equals("Desired Salary")){ j16=true;count++;}
	        	 if(fields.get(x).equals("SBU")){ j17=true;count++;}if(fields.get(x).equals("Referred By Email")){ j18=true;count++;}
	        	 if(fields.get(x).equals("Referred By")){ j19=true;count++;}
	        	 
	         	 }if(fields.size()>19){
	         		 return 10;
	         	 }
	         	 else if(count==19){
	         		 System.out.println(extra);
	        		return count;
	        	 }
	         	 
	        	 System.out.println("verifying");
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
	 
public void checking(){
	 try{

		 heading.setBounds(10,20,900,30); 
    	 heading.setText("These fields are not present");
    	 heading.setHorizontalAlignment(JLabel.CENTER);
    	 if(j1){l1.setVisible(false);}if(j2){l2.setVisible(false);}if(j3){l3.setVisible(false);}
    	 if(j4){l4.setVisible(false);}if(j5){l5.setVisible(false);}if(j6){l6.setVisible(false);}
    	 if(j7){l7.setVisible(false);}if(j8){l8.setVisible(false);}if(j9){l9.setVisible(false);}
    	 if(j10){l10.setVisible(false);}if(j11){l11.setVisible(false);}if(j12){l12.setVisible(false);}
    	 if(j13){l13.setVisible(false);}if(j14){l14.setVisible(false);}if(j15){l15.setVisible(false);}
    	 if(j16){l16.setVisible(false);}if(j17){l17.setVisible(false);}if(j18){l18.setVisible(false);}
    	 if(j19){l19.setVisible(false);}
    	 
		 if(j1&&j2&&j3&&j4&&j5&&j6&&j7&&j8&&j9&&j10&&j11&&j12&&j13&&j14&&j15&&j16&&j17&&j18&&j19){
			 showMiss();b2.setVisible(false);
			 b1.setVisible(false);
			
		 }
    		 b1.setVisible(false);
    		 b2 = new JButton("Show Extra Fields");
 	        b2.setBounds(330,270,230,30);
 	        b2.setBackground(new Color(255,82,82));
 	        b2.setForeground(new Color(255,255,255));
 	        b2.addActionListener(new ActionListener() {
		        public void actionPerformed(ActionEvent e) {
		        	showMiss();
			       }			
			     });
 	        d.add(b2);d.addKeyListener(this);
    	 }catch(NullPointerException er){}
}

public void showMiss(){
	l1.setText("");l2.setText("");l3.setText("");l4.setText("");l5.setText("");
	l6.setText("");l7.setText("");l8.setText("");l9.setText("");l10.setText("");
	l11.setText("");l12.setText("");l13.setText("");l14.setText("");l15.setText("");
	l16.setText("");l17.setText("");l18.setText("");l19.setText("");
		 heading.setText("These are the extra fields present in the Master Tracker File");
		 fields.remove("Referred By");fields.remove("Title");fields.remove("Candidate Full Name");fields.remove("Candidate ID");
		 fields.remove("Candidate Email");fields.remove("REQ #");fields.remove("Applied Date (WEB)");fields.remove("Applied Date (WEB/MCH)");
		 fields.remove("Business Unit (Req More)");fields.remove("Business Unit (Hierarchy)");fields.remove("Candidate Phone Number");fields.remove("Candidate Source");
		 fields.remove("Candidate Skills");fields.remove("Cell Phone");fields.remove("Cell telephone");fields.remove("Current Salary Rate");
		 fields.remove("Desired Salary");fields.remove("SBU");fields.remove("Referred By Email");
		 try{
		 l1.setVisible(true);l1.setText(fields.get(0).toString());
		 l2.setText(fields.get(1).toString());l2.setVisible(true);l2.setBounds(170,80,250,30);
		 l3.setText(fields.get(2).toString());l3.setVisible(true);
		 l4.setText(fields.get(3).toString());l4.setVisible(true);
		 l5.setText(fields.get(4).toString());l5.setVisible(true);
		 l6.setText(fields.get(5).toString());l6.setVisible(true);
		 l7.setText(fields.get(6).toString());l7.setVisible(true);
		 l8.setText(fields.get(7).toString());l8.setVisible(true);
		 l9.setText(fields.get(8).toString());l9.setVisible(true);
		 l10.setText(fields.get(9).toString());l10.setVisible(true);
		 l11.setText(fields.get(10).toString());l11.setVisible(true);
		 l12.setText(fields.get(11).toString());l12.setVisible(true);
		 l13.setText(fields.get(12).toString());l13.setVisible(true);
		 l14.setText(fields.get(13).toString());l14.setVisible(true);
		 l15.setText(fields.get(14).toString());l15.setVisible(true);
		 }catch(IndexOutOfBoundsException|NullPointerException ad){}
	 
}
	 public static void main(String...args){
		 String x ="(CGI) Requisition Applicants (5).xls";
		 new helpMaster().helpMasters(x);

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
