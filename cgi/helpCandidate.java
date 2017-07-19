package cgi;
import javax.swing.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.awt.*;  
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;  

public class helpCandidate implements KeyListener {
	 private static JDialog d;  
	 JLabel heading,l1,l2,l3,l4,l5,l6,l7,l8,l9,l10,l11,l12,l13,l14,l15,l16,l17,l18,l19,l20,l21,l22,l23,l24;
	 Boolean j1=false,j2=false,j3=false,j4=false,j5=false,j6=false,j7=false,j8=false,j9=false,j10=false,j11=false,j12=false,j13=false,j14=false,j15=false,j16=false,j17=false,j18=false,j19=false,j20=false,j21=false,j22=false,j23=false,j24=false;
	 java.util.List<String> fields = new ArrayList<String>();
	 JButton b1,b2;
	 int extra = 0;
public void	 justhelpCandidate(){
	       
		  JFrame f= new JFrame();
		  
	        d = new JDialog(f , "Master Referral Validation Automator", true);  
	        d.setLayout(null);  
	        heading = new JLabel("Please Check that the Candidate Referral Input File has these fields in any order : ");
	        heading.setBounds(10,20,900,30); 
	        heading.setFont(new Font("Tahoma",Font.PLAIN,20));
	        heading.setForeground(Color.BLUE);
	        l1 = new JLabel(" Candidate First Name");
	        l1.setBounds(10,80,200,30);
	        l1.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2 = new JLabel(" Candidate Last Name");
	        l2.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2.setBounds(210,80,250,30);
	        l3 = new JLabel(" Candidate Email");
	        l3.setBounds(410,80,150,30);
	        l3.setFont(new Font("Tahoma",Font.BOLD,15));
	        l4 = new JLabel(" Candidate Source");
	        l4.setBounds(580,80,150,30);
	        l4.setFont(new Font("Tahoma",Font.BOLD,15));
	        l5 = new JLabel(" Referral Name");
	        l5.setBounds(750,80,150,30);
	        l5.setFont(new Font("Tahoma",Font.BOLD,15));
	        l6 = new JLabel(" Referral Email");
	        l6.setBounds(10,120,250,30);
	        l6.setFont(new Font("Tahoma",Font.BOLD,15));
	        l7 = new JLabel(" JOB ID");
	        l7.setBounds(210,120,250,30);
	        l7.setFont(new Font("Tahoma",Font.BOLD,15));
	        l8 = new JLabel(" Job Title");
	        l8.setBounds(410,120,250,30);
	        l8.setFont(new Font("Tahoma",Font.BOLD,15));
	        l9 = new JLabel(" Candidate Stage");
	        l9.setBounds(580,120,250,30);
	        l9.setFont(new Font("Tahoma",Font.BOLD,15));
	        l10 = new JLabel(" Application Status");
	        l10.setBounds(750,120,250,30);
	        l10.setFont(new Font("Tahoma",Font.BOLD,15));
	        l11 = new JLabel(" Application Date");
	        l11.setFont(new Font("Tahoma",Font.BOLD,15));
	        l11.setBounds(10,160,250,30);
	        l12 = new JLabel(" Ref. Survey Status");
	        l12.setFont(new Font("Tahoma",Font.BOLD,15));
	        l12.setBounds(210,160,250,30);
	        l13 = new JLabel(" Date Survey Taken");
	        l13.setFont(new Font("Tahoma",Font.BOLD,15));
	        l13.setBounds(410,160,250,30);
	        l14 = new JLabel(" Date Survey Invite Sent");
	        l14.setFont(new Font("Tahoma",Font.BOLD,15));
	        l14.setBounds(580,160,250,30);
	        l15 = new JLabel(" Candidate Enter Date");
	        l15.setFont(new Font("Tahoma",Font.BOLD,15));
	        l15.setBounds(10,200,250,30);
	        l16 = new JLabel(" Referral Type");
	        l16.setFont(new Font("Tahoma",Font.BOLD,15));
	        l16.setBounds(210,200,250,30);
	        l17 = new JLabel(" Placement Date");
	        l17.setFont(new Font("Tahoma",Font.BOLD,15));
	        l17.setBounds(410,200,250,30);
	        l18 = new JLabel(" Start Date");
	        l18.setFont(new Font("Tahoma",Font.BOLD,15));
	        l18.setBounds(580,200,250,30);
	        l19 = new JLabel(" Business Unit");
	        l19.setFont(new Font("Tahoma",Font.BOLD,15));
	        l19.setBounds(7500,200,250,30);
	        l20 = new JLabel(" CandidateID");
	        l20.setFont(new Font("Tahoma",Font.BOLD,15));
	        l20.setBounds(10,240,250,30);
	        l21 = new JLabel(" Last Activity Date");
	        l21.setFont(new Font("Tahoma",Font.BOLD,15));
	        l21.setBounds(210,240,250,30);
	        l22 = new JLabel(" Referral ID");
	        l22.setFont(new Font("Tahoma",Font.BOLD,15));
	        l22.setBounds(410,240,250,30);
	        l23 = new JLabel(" SVP/VP");
	        l23.setFont(new Font("Tahoma",Font.BOLD,15));
	        l23.setBounds(580,240,250,30);
	        l24 = new JLabel(" VP/Director");
	        l24.setFont(new Font("Tahoma",Font.BOLD,15));
	        l24.setBounds(750,240,250,30);
	        JLabel txt = new JLabel("Press any key to Exit...");
	        txt.setBounds(370,280,250,30);
	        txt.setForeground(Color.RED);
	        txt.setFont(new Font("Tahoma",Font.PLAIN,14));
	        
	        
	        
	        
	        d.addKeyListener(this);  
	        d.add(heading);d.add(txt);
	        d.add(l1);d.add(l2);d.add(l3);d.add(l4);d.add(l5);d.add(l6);d.add(l7);d.add(l8);d.add(l9);d.add(l10);
	        d.add(l11);d.add(l12);d.add(l13);d.add(l14);d.add(l15);d.add(l16);d.add(l17);d.add(l18);d.add(l19);d.add(l20);
	        d.add(l21);d.add(l22);d.add(l23);d.add(l24);
	        d.setSize(940,360);
	        d.setLocation(250,30);
	      
	        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        d.setVisible(true);  
	       
	        
	 }
	 
public void	 helpCandidates(String input2){
    
	  JFrame f= new JFrame();
	  
      d = new JDialog(f , "Master Referral Validation Automator", true);  
      d.setLayout(null);  
      heading = new JLabel("Please Check that the Candidate Referral Input File has these fields in any order : ");
      heading.setBounds(10,20,900,30); 
      heading.setFont(new Font("Tahoma",Font.PLAIN,20));
      heading.setForeground(Color.BLUE);
      l1 = new JLabel(" Candidate First Name");
      l1.setBounds(10,80,200,30);
      l1.setFont(new Font("Tahoma",Font.BOLD,15));
      l2 = new JLabel(" Candidate Last Name");
      l2.setFont(new Font("Tahoma",Font.BOLD,15));
      l2.setBounds(210,80,250,30);
      l3 = new JLabel(" Candidate Email");
      l3.setBounds(410,80,150,30);
      l3.setFont(new Font("Tahoma",Font.BOLD,15));
      l4 = new JLabel(" Candidate Source");
      l4.setBounds(580,80,150,30);
      l4.setFont(new Font("Tahoma",Font.BOLD,15));
      l5 = new JLabel(" Referral Name");
      l5.setBounds(750,80,150,30);
      l5.setFont(new Font("Tahoma",Font.BOLD,15));
      l6 = new JLabel(" Referral Email");
      l6.setBounds(10,120,250,30);
      l6.setFont(new Font("Tahoma",Font.BOLD,15));
      l7 = new JLabel(" JOB ID");
      l7.setBounds(210,120,250,30);
      l7.setFont(new Font("Tahoma",Font.BOLD,15));
      l8 = new JLabel(" Job Title");
      l8.setBounds(410,120,250,30);
      l8.setFont(new Font("Tahoma",Font.BOLD,15));
      l9 = new JLabel(" Candidate Stage");
      l9.setBounds(580,120,250,30);
      l9.setFont(new Font("Tahoma",Font.BOLD,15));
      l10 = new JLabel(" Application Status");
      l10.setBounds(750,120,250,30);
      l10.setFont(new Font("Tahoma",Font.BOLD,15));
      l11 = new JLabel(" Application Date");
      l11.setFont(new Font("Tahoma",Font.BOLD,15));
      l11.setBounds(10,160,250,30);
      l12 = new JLabel(" Ref. Survey Status");
      l12.setFont(new Font("Tahoma",Font.BOLD,15));
      l12.setBounds(210,160,250,30);
      l13 = new JLabel(" Date Survey Taken");
      l13.setFont(new Font("Tahoma",Font.BOLD,15));
      l13.setBounds(410,160,250,30);
      l14 = new JLabel(" Date Survey Invite Sent");
      l14.setFont(new Font("Tahoma",Font.BOLD,15));
      l14.setBounds(580,160,250,30);
      l15 = new JLabel(" Candidate Enter Date");
      l15.setFont(new Font("Tahoma",Font.BOLD,15));
      l15.setBounds(10,200,250,30);
      l16 = new JLabel(" Referral Type");
      l16.setFont(new Font("Tahoma",Font.BOLD,15));
      l16.setBounds(210,200,250,30);
      l17 = new JLabel(" Placement Date");
      l17.setFont(new Font("Tahoma",Font.BOLD,15));
      l17.setBounds(410,200,250,30);
      l18 = new JLabel(" Start Date");
      l18.setFont(new Font("Tahoma",Font.BOLD,15));
      l18.setBounds(580,200,250,30);
      l19 = new JLabel(" Business Unit");
      l19.setFont(new Font("Tahoma",Font.BOLD,15));
      l19.setBounds(7500,200,250,30);
      l20 = new JLabel(" CandidateID");
      l20.setFont(new Font("Tahoma",Font.BOLD,15));
      l20.setBounds(10,240,250,30);
      l21 = new JLabel(" Last Activity Date");
      l21.setFont(new Font("Tahoma",Font.BOLD,15));
      l21.setBounds(210,240,250,30);
      l22 = new JLabel(" Referral ID");
      l22.setFont(new Font("Tahoma",Font.BOLD,15));
      l22.setBounds(410,240,250,30);
      l23 = new JLabel(" SVP/VP");
      l23.setFont(new Font("Tahoma",Font.BOLD,15));
      l23.setBounds(580,240,250,30);
      l24 = new JLabel(" VP/Director");
      l24.setFont(new Font("Tahoma",Font.BOLD,15));
      l24.setBounds(750,240,250,30);
      JLabel txt = new JLabel("Press any key to Exit...");
      checkFormat(input2);
      txt.setBounds(370,330,250,30);
      txt.setForeground(Color.RED);
      txt.setFont(new Font("Tahoma",Font.PLAIN,14));
      b1 = new JButton("Show Error with Fields");
      b1.setBounds(330,300,230,30);
      b1.setForeground(new Color(255,255,255));
      b1.setBackground(new Color(255,82,82));
      b1.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
		        checking();}			
		     });
      
      d.addKeyListener(this);d.setResizable(false);
      d.add(heading);d.add(txt);d.add(b1);
      d.add(l1);d.add(l2);d.add(l3);d.add(l4);d.add(l5);d.add(l6);d.add(l7);d.add(l8);d.add(l9);d.add(l10);
      d.add(l11);d.add(l12);d.add(l13);d.add(l14);d.add(l15);d.add(l16);d.add(l17);d.add(l18);d.add(l19);d.add(l20);
      d.add(l21);d.add(l22);d.add(l23);d.add(l24);
      d.setSize(930,400);
      d.setLocation(250,30);
    
      f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
      d.setVisible(true);  
     
      
}

public int checkFormat(String input1){
		 try{
			
			 int count2=0;
			 if(input1.contains(".xls")||input1.contains(".xlsx")){
		     FileInputStream inputStream = new FileInputStream(new File(input1));
		     Workbook workbook = WorkbookFactory.create(inputStream);
		     Sheet sheet = workbook.getSheetAt(0);
		         Row ttya = sheet.getRow(5);
		         int colcount2 = ttya.getLastCellNum();
		         for(int i =0;i<colcount2;i++){
		        	 Cell cell = ttya.getCell(i);
		        	 if(cell==null){
		        		 continue;
		        	 }
		        	 fields.add(cell.getStringCellValue());
		         }for(int x = 0;x<fields.size();x++){
		        	 if(fields.get(x).equals("Candidate First Name")){ j1=true;count2++;}if(fields.get(x).equals("Candidate Last Name")){ j2=true;count2++;}
		        	 if(fields.get(x).equals("Candidate Email")){ j3=true;count2++;}if(fields.get(x).equals("Candidate Source")){ j4=true;count2++;}
		        	 if(fields.get(x).equals("Referral Name")){ j5=true;count2++;}if(fields.get(x).equals("Referral Email")){ j6=true;count2++;}
	        	     if(fields.get(x).equals("Job ID")){ j7=true;count2++;}if(fields.get(x).equals("Job Title")){ j8=true;count2++;}
		        	 if(fields.get(x).equals("Candidate Stage")){ j9=true;count2++;}if(fields.get(x).equals("Application Status")){ j10=true;count2++;}
		        	 if(fields.get(x).equals("Application Date")){ j11=true;count2++;}if(fields.get(x).equals("Ref. Survey Status")){ j12=true;count2++;}
		        	 if(fields.get(x).equals("Date Survey Taken")){ j13=true;count2++;}if(fields.get(x).equals("Date Survey Invite Sent")){ j14=true;count2++;}
		        	 if(fields.get(x).equals("Candidate Enter Date")){ j15=true;count2++;}if(fields.get(x).equals("Referral Type")){ j16=true;count2++;}
		        	 if(fields.get(x).equals("Placement Date")){ j17=true;count2++;}if(fields.get(x).equals("Start Date")){ j18=true;count2++;}
		        	 if(fields.get(x).equals("Business Unit")) {j19=true;count2++;}if(fields.get(x).equals("CandidateID")) {j20=true;count2++;}
		        	 if(fields.get(x).equals("Last Activity Date")) {j21=true;count2++;}if(fields.get(x).equals("Referral ID")) {j22=true;count2++;}
		        	 if(fields.get(x).equals("SVP/VP")) {j23=true;count2++;}if(fields.get(x).equals("VP/Director")) {j24=true;count2++;}
		         }
		         if(fields.size()>24){
		        	 return 19;
		         }
		        	 if(count2==24){
		        		return count2;
		        	 }
		        	 System.out.println("verifying");
		        	 return count2;
			 } 
			 else{
				 throw new IOException();
			 }
		 }catch (IOException  | EncryptedDocumentException | IndexOutOfBoundsException
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
   	 if(j19){l19.setVisible(false);}if(j20){l20.setVisible(false);}if(j21){l21.setVisible(false);}
   	if(j22){l22.setVisible(false);}if(j23){l23.setVisible(false);}if(j24){l24.setVisible(false);}
   	if(j1&&j2&&j3&&j4&&j5&&j6&&j7&&j8&&j9&&j10&&j11&&j12&&j13&&j14&&j15&&j16&&j17&&j18&&j19&&j20&&j21&&j22&&j23&&j24){
   		showMiss();		 b2.setVisible(false);
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
    d.add(b2);
   	 }catch(NullPointerException er){}
}

public void showMiss(){
	l1.setText("");l2.setText("");l3.setText("");l4.setText("");l5.setText("");
	l6.setText("");l7.setText("");l8.setText("");l9.setText("");l10.setText("");
	l11.setText("");l12.setText("");l13.setText("");l14.setText("");l15.setText("");
	l16.setText("");l17.setText("");l18.setText("");l19.setText("");l20.setText("");l21.setText("");
	l22.setText("");l23.setText("");l24.setText("");
  	 heading.setText("These are the extra fields present in the Candidate Referral File");
  	 fields.remove("Candidate First Name");fields.remove("Candidate Last Name");fields.remove("Candidate Email");fields.remove("Candidate Source");
	 fields.remove("Referral Name");fields.remove("Referral Email");fields.remove("Job ID");fields.remove("Job Title");
	 fields.remove("Candidate Stage");fields.remove("Application Status");fields.remove("Application Date");fields.remove("Ref. Survey Status");
	 fields.remove("Date Survey Taken");fields.remove("Date Survey Invite Sent");fields.remove("Candidate Enter Date");fields.remove("Referral Type");
	 fields.remove("Placement Date");fields.remove("Start Date");fields.remove("Business Unit");fields.remove("Business Unit");fields.remove("CandidateID");
	 fields.remove("Last Activity Date");fields.remove("Referral ID");fields.remove("SVP/VP");fields.remove("VP/Director");
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
		 String x ="Copy of Candidate Referrals (Generic).xls";
		 new helpCandidate().helpCandidates(x);
		 //checkFormat(x);
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
