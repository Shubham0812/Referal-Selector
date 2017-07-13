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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.monitorjbl.xlsx.StreamingReader;

import cgi.AMSdump.DataStorer;  

public class helpAms implements KeyListener {
	 private static JDialog d2; 
	 java.util.List<String> fields = new ArrayList<String>();
	 JLabel heading,l1,l2,l3,l4,l5,l6,l7;
		java.util.List<DataStorer> data = new ArrayList<DataStorer>();
	 public void helpAmsa(){
		  JFrame fae= new JFrame();
		  
	        d2 = new JDialog(fae , "Master Referral Validation Automator", true);  
	        d2.setLayout(null);  
	        
	        heading = new JLabel("Please Check that the Ams Dump File's Field Matches with the Following : ");
	        heading.setBounds(10,20,900,30); 
	        heading.setFont(new Font("Tahoma",Font.PLAIN,20));
	        heading.setForeground(Color.BLUE);
	        l1 = new JLabel("1. mobile1");
	        l1.setBounds(10,80,150,30);
	        l1.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2 = new JLabel("2. email1");
	        l2.setFont(new Font("Tahoma",Font.BOLD,15));
	        l2.setBounds(120,80,250,30);
	        l3 = new JLabel("3. candidate_id");
	        l3.setBounds(220,80,150,30);
	        l3.setFont(new Font("Tahoma",Font.BOLD,15));
	        l4 = new JLabel("4. PanNo");
	        l4.setBounds(370,80,150,30);
	        l4.setFont(new Font("Tahoma",Font.BOLD,15));
	        l5 = new JLabel("5. SOURCE_NAME");
	        l5.setBounds(470,80,150,30);
	        l5.setFont(new Font("Tahoma",Font.BOLD,15));
	        l6 = new JLabel("6. CurrentStage");
	        l6.setBounds(630,80,250,30);
	        l6.setFont(new Font("Tahoma",Font.BOLD,15));
	        l7 = new JLabel("7. current_status");
	        l7.setBounds(760,80,250,30);
	        l7.setFont(new Font("Tahoma",Font.BOLD,15));
	        JLabel txt = new JLabel("Press any key to Exit...");
	        txt.setBounds(370,250,250,30);
	        txt.setForeground(Color.RED);
	        txt.setFont(new Font("Tahoma",Font.PLAIN,14));
	        
	        
	        
	        
	        d2.addKeyListener(this);  
	        d2.add(heading);d2.add(txt);
	        d2.add(l1);d2.add(l2);d2.add(l3);d2.add(l4);d2.add(l5);d2.add(l6);d2.add(l7);
	        d2.setSize(940,340);
	        d2.setLocation(250,30);
	      
	        fae.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        d2.setVisible(true);  
	       
	        
}
	 
	 public void storeIntoList(String output2){
			File is = new File(output2);
			Workbook workbook = StreamingReader.builder()
								.rowCacheSize(100)
								.bufferSize(1096)
								.open(is);
			int cols = 0;int rows = 0;
			  for (Sheet sheet : workbook){
			    	for (Row r : sheet) {
			    		for (Cell c : r) {
			    			if(rows==1){
			    				return;
			    			}
			    			fields.add(c.getStringCellValue());	
			    	}
			    		rows++;
			    	}

		}
	 }
	 
	 
public int checkFormat(String input2) {
	 try{
		 int count3=0;
		 
		 if(input2.contains(".xls")||input2.contains(".xlsx")){
			 storeIntoList(input2);
			 if(fields.get(0).equals("mobile1")){ count3++;}if(fields.get(1).equals("email1")){ count3++;}
        	 if(fields.get(2).equals("candidate_id")){ count3++;}if(fields.get(3).equals("PanNo")){ count3++;}
        	 if(fields.get(4).equals("SOURCE_NAME")){ count3++;}if(fields.get(5).equals("CurrentStage")){ count3++;}
        	 if(fields.get(6).equals("current_status")){ count3++;}
        	 
        	 if(count3==7){
        		 int ans = count3;
        		 count3 = 0;
        		 System.out.println("YES!" + count3);
        		return ans;
        	 }
        	 count3 = 0;
        	 return count3;
	 }
	 else{
		 throw new IOException();
	 }
	 }catch (EncryptedDocumentException | IOException | IndexOutOfBoundsException ex ) {
     JOptionPane.showMessageDialog(null, "Error : Invalid File Selected");
     return -1;
	 }
}
	 
	 public static void main(String...args) throws InvalidFormatException, IOException{
		 new helpAms();
		 String x ="AMS_Dump_Datax.xlsx";
	 }

	@Override
	public void keyPressed(KeyEvent e) {
		d2.dispose();
	}

	@Override
	public void keyReleased(KeyEvent arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void keyTyped(KeyEvent arg0) {
		}
		
}
