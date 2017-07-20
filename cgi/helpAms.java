/* Automation of Member Referral Process (June-2017) 
 * Author - Shubham Kumar Singh
 * Email - singh.shubham0812@gmail.com
 * College - Nitte Meenakshi Institute of Technology, Bangalore 
 */
//program to check the format of the Ams Dump
package cgi;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import com.monitorjbl.xlsx.StreamingReader;
import cgi.AMSdump.DataStorer;  

public class helpAms implements KeyListener {
private static JDialog d2; 
java.util.List<String> fields = new ArrayList<String>();
JLabel heading,l1,l2,l3,l4,l5,l6,l7;int extra;
boolean j1=false,j2=false,j3=false,j4=false,j5=false,j6=false,j7=false;
JButton b1,b2;
java.util.List<DataStorer> data = new ArrayList<DataStorer>();

public void justHelp(){
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

public void helpAmss(String input){
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
    b1 = new JButton("Show Error with Fields");
    b1.setBounds(330,270,230,30);
    b1.setForeground(new Color(255,255,255));
    JLabel txt = new JLabel("Press any key to Exit...");
    txt.setBounds(375,300,250,30);;
    txt.setForeground(Color.RED);
    txt.setFont(new Font("Tahoma",Font.PLAIN,14));
    checkFormat(input);
    b1.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent e) {
    	checking();}			
});       
    b1.setBackground(new Color(255,82,82));
    d2.addKeyListener(this);  
    d2.add(heading);d2.add(txt);
    d2.add(l1);d2.add(l2);d2.add(l3);d2.add(l4);d2.add(l5);d2.add(l6);d2.add(l7);d2.add(b1);
    d2.setSize(930,380);
    d2.setResizable(false);
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
	int rows = 0;
	for (Sheet sheet : workbook){
		for (Row r : sheet) {
			for (Cell c : r) {
				if(rows==1){
					return;}
				fields.add(c.getStringCellValue());	}
			rows++;}
	}
 }
 
	 
public int checkFormat(String input2) {
	 try{
		 int count3=0;
		 if(input2.contains(".xls")||input2.contains(".xlsx")){
			 storeIntoList(input2);
			 extra = (fields.size()-19);
			 for(int x = 0;x<fields.size();x++){
				 if(fields.get(x).equals("mobile1")){ count3++;j1=true;}if(fields.get(x).equals("email1")){ j2=true;count3++;}
				 if(fields.get(x).equals("candidate_id")){ j3=true;count3++;}if(fields.get(x).equals("PanNo")){j4=true; count3++;}
				 if(fields.get(x).equals("SOURCE_NAME")){j5=true; count3++;}if(fields.get(x).equals("CurrentStage")){j6=true; count3++;}
				 if(fields.get(x).equals("current_status")){ j7=true;count3++;}
			 }
			 if(fields.size()>7){
				 return 11;
			 }
			 else if(count3==7){
        		return count3;
        	 }
			 return count3;
		 }
		 else{
		 throw new IOException();}
	 }catch (Exception e ) {
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
		 if(j7){l7.setVisible(false);}
		 if(j1&&j2&&j3&&j4&&j5&&j6&&j7){
			 showMiss(); b2.setVisible(false);
			 b1.setVisible(false);}
		 b1.setVisible(false);
		 b2 = new JButton("Show Extra Fields");
		 b2.setBounds(330,270,230,30);
		 b2.setBackground(new Color(255,82,82));
		 b2.setForeground(new Color(255,255,255));
	     b2.addActionListener(new ActionListener() {
         public void actionPerformed(ActionEvent e) {
        	showMiss();}			
	     });
	     d2.add(b2);
   	 }catch(NullPointerException er){}
}

public void showMiss(){
	l1.setText("");l2.setText("");l3.setText("");l4.setText("");l5.setText("");
	l6.setText("");l7.setText("");
	heading.setText("These are the extra fields present in the Ams Dump File");
	fields.remove("mobile1");fields.remove("email1");fields.remove("candidate_id");fields.remove("PanNo");
	fields.remove("SOURCE_NAME");fields.remove("CurrentStage");fields.remove("current_status");
	try{
		l1.setVisible(true);l1.setText(fields.get(0).toString());
		l2.setText(fields.get(1).toString());l2.setVisible(true);l2.setBounds(170,80,250,30);
		l3.setText(fields.get(2).toString());l3.setVisible(true);
		l4.setText(fields.get(3).toString());l4.setVisible(true);
		l5.setText(fields.get(4).toString());l5.setVisible(true);
		l6.setText(fields.get(5).toString());l6.setVisible(true);
		l7.setText(fields.get(6).toString());l7.setVisible(true);
	 }catch(IndexOutOfBoundsException|NullPointerException ad){}
}

public void keyPressed(KeyEvent e) {d2.dispose();}
public void keyReleased(KeyEvent arg0) {}
public void keyTyped(KeyEvent arg0) {}

//main method
public static void main(String...args) throws InvalidFormatException, IOException{
		 new helpAms();}



		
}
