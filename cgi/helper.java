/* Automation of Member Referral Process (June-2017) 
 * Author - Shubham Kumar Singh
 * Email - singh.shubham0812@gmail.com
 * College - Nitte Meenakshi Institute of Technology, Bangalore 
 */

//helper class to tell more about the application
package cgi;

import java.awt.*;
import java.awt.event.*;
import java.util.ArrayList;
import javax.swing.*;
import javax.swing.border.Border;
import cgi.AMSdump.DataStorer;  
public class helper implements KeyListener {  
private static JDialog d2; 
static int count3=0;
JLabel heading,l1,l2,l3,l4,l5,l6,l7;
static java.util.List<String> fields = new ArrayList<String>();
java.util.List<DataStorer> data = new ArrayList<DataStorer>();
helper(){
	JFrame fae= new JFrame();
	Border blackline = BorderFactory.createLineBorder(Color.black);
    d2 = new JDialog(fae , "Master Referral Validation Automator--HELP", true);  
    d2.setLayout(null);  
    heading = new JLabel(" HELP");
    heading.setBounds(10,20,80,40); 
    heading.setFont(new Font("Tahoma",Font.BOLD,26));
    heading.setForeground(Color.BLUE);
    heading.setBorder(blackline);
    JTextArea area = new JTextArea();
    area.setText("");
    Color color = UIManager.getColor ( "Frame.background" );
    area.setBackground(color);
    area.setBounds(200,60,567,77);
    area.setText("The application's purpose is to Automate the Member Referral Process\nIt contains different modules which automates the different parts of the\nprocess.");
    area.setFont(new Font("Times New Roman",Font.BOLD,18));area.setFocusable(false);
    area.setEditable(false);
    UIManager.put("TabbedPane.background",new Color(107,210,219));
    JTabbedPane tp=new JTabbedPane();  
    tp.setBounds(170,170,600,300);  
    JLabel txt = new JLabel("It is Recommended to Check the format of the Input Files using the options from the Help Menu!");
    txt.setBounds(50,458,850,30);
    txt.setForeground(Color.RED);
    txt.setFont(new Font("Tahoma",Font.BOLD,17));
    //panels
    JPanel p1=new JPanel();  
    JPanel p2 = new JPanel();
    JPanel p3 = new JPanel();
    JPanel p4 = new JPanel();
    p1.setLayout(new FlowLayout());
    p2.setLayout(new FlowLayout());
    p3.setLayout(new FlowLayout());
    p4.setLayout(new FlowLayout());
    JTextArea t1 = new JTextArea();
    t1.setBackground(color);
    t1.setFont(new Font("Tahoma",Font.PLAIN,18));
    t1.setEditable(false);
    t1.setText("This module requires two Files : \n 1. The Master Tracker File \n 2. The Candidate Referral File\n\nThe Module performs a Vlookup from these two files taking \nValidation Index( Concatenation of REQ# and Candidate ID)\nas reference and gets the Referred By Name & Referred By Email\nin the output file which also contains the MR data and these 2 fields.\n\n");
    p1.add(t1);
    JLabel l2 = new JLabel("Output File Name : VLookupOutputs.xlsx  ");
    l2.setFont(new Font("Times New Roman",Font.BOLD,18));
    p1.add(l2);
    tp.add("Module 1: VLookup",p1);  
    JTextArea t2 = new JTextArea();
    t2.setBackground(color);
    t2.setFont(new Font("Tahoma",Font.PLAIN,18));
    t2.setEditable(false);
    t2.setText("This module requires two Files : \n 1. The VLookup Output File \n 2. The AMS Dump File\n\nThe Module performs a Vlookup from these two files taking Mobile#\nand Email ID as reference and does a Duplicacy Check,\nand also gets the Source, Current Stage and Current Status\nin the output file which also contains the MR file data and\nthese fields.\n");
    p2.add(t2);
    JLabel l3 = new JLabel("Output File Name : AmsDumpOutput.xlsx  ");
    l3.setFont(new Font("Times New Roman",Font.BOLD,18));
    p2.add(l3);
    tp.add("Module 2: Check From AMS Dump",p2);
    JTextArea t3 = new JTextArea();
    t3.setBackground(color);
    t3.setFont(new Font("Tahoma",Font.PLAIN,18));
    t3.setEditable(false);
    t3.setText("This module requires only one File : \n 1. The AmsDumpOutput File\n\nThe Module gets the Communcation Channels for the all the record in the\nAmsOutput file by Checking whether they are unique, duplicate or Tocheck.\nIn case of unique One channel is given and in the other cases the sources\nare checked to see whether they are RA-MR-P,  RA-P or  ER.\nIn this cases the current stage and current status are checked\nand corresponding Mails are decided for them.\n");
    p3.add(t3);
    JLabel l4 = new JLabel("Output File Name : AmsCommMails.xlsx  ");
    l4.setFont(new Font("Times New Roman",Font.BOLD,18));
    p3.add(l4);
    tp.add("Module 3: Get Communication Mails",p3);
    JTextArea t4 = new JTextArea();
    t4.setBackground(color);
    t4.setFont(new Font("Tahoma",Font.PLAIN,18));
    t4.setEditable(false);
    t4.setText("This module requires three Files : \n1. The Master Tracker File\n2.The Candidate Referral File\n3.The Ams Dump File\n\nThe Module performs all the processes of MR Validation one by one.\n1.Vlookup is performed using the Master Tracker and the\nCandidate Referral file, and gets Referred By Name & Email \n\n2. The Duplicacy check is performed with the help of Ams Dump File\nand duplicacy check is done for candidates along with getting the \nSource, Current Stage & Status.\n\n3. This module gets the Communcation mailers for different cases for\nthe candidates, in case of UNIQUE \"comm1\" is alloted\nIn case of DUPLICATE with no Source \"Comm 2a\\2b\" is alloted.\nIn case of duplicates with either RA-MR-P,  RA-P or  ER\nthen different Comms are alloted based on current Source & Status.");
    t2.setFocusable(false);t1.setFocusable(false);t3.setFocusable(false);t4.setFocusable(false);
    p4.add(t4);
    d2.setResizable(false);
    tp.add(new JScrollPane(p4),"Module 4: Execute All At Once");
    d2.addKeyListener(this);  
    d2.add(heading);d2.add(txt);
    d2.setSize(940,540);d2.add(area);d2.add(tp);
    d2.setLocation(250,30);
    d2.addKeyListener(this);
    fae.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    d2.setVisible(true);   
}
	public void keyPressed(KeyEvent e) {d2.dispose();}
	public void keyReleased(KeyEvent arg0) {}
	public void keyTyped(KeyEvent arg0) {}	
	
//caling the main method
	public static void main(String...strings){
		new helper();}
}  