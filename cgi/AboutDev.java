/* Automation of Member Referral Process (June-2017) 
 * Author - Shubham Kumar Singh
 * Email - singh.shubham0812@gmail.com
 * College - Nitte Meenakshi Institute of Technology, Bangalore 
 */

//this class shows the information about the application and the developer
package cgi;

import java.awt.Color;
import java.awt.Font;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.BorderFactory;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import javax.swing.border.Border;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import cgi.AMSdump.DataStorer;  

public class AboutDev implements KeyListener {
private static JDialog d2; 
static int count3=0;
JLabel heading,l1,l2,l3,l4,subheading;
static java.util.List<String> fields = new ArrayList<String>();
java.util.List<DataStorer> data = new ArrayList<DataStorer>();
AboutDev(){
	JFrame fae= new JFrame();
	d2 = new JDialog(fae , "Member Referral Validation Automator v1.02 --ABOUT", true);  
    d2.setLayout(null);  
    Border blackline = BorderFactory.createLineBorder(Color.black);
    heading = new JLabel("Member Referral Validation Automator v1.02",SwingConstants.CENTER);
    heading.setBounds(160,20,500,30); 
    heading.setForeground(Color.CYAN);
    heading.setFont(new Font("Tahoma",Font.PLAIN,20));
    heading.setForeground(Color.BLUE);
    subheading = new JLabel("About",SwingConstants.CENTER);
    subheading.setBounds(310,55,150,30); 
    subheading.setForeground(Color.RED);
    subheading.setBorder(blackline);
    subheading.setFont(new Font("Tahoma",Font.PLAIN,20));
    l1 = new JLabel("This Application automates the Member Referral Process by obtaining the Communication Mailers.");
    l1.setBounds(10,89,950,30);
    l1.setFont(new Font("Tahoma",Font.BOLD,15));
    l2 = new JLabel("Java along with Swing was used to developed the application along with Apache POI.");
    l2.setFont(new Font("Tahoma",Font.BOLD,15));
    l2.setBounds(10,130,750,30);
    l3 = new JLabel(" Developed by : Shubham Kumar Singh    |    Email-ID : singh.shubham0812@gmail.com");
    l3.setBounds(10,170,660,30);
    l3.setFont(new Font("Tahoma",Font.BOLD,15));
    l4 = new JLabel("If there are any bugs found send a mail to the email specified above.");
    l4.setBounds(10,210,750,30);
    l4.setFont(new Font("Tahoma",Font.BOLD,15));
    l3.setBorder(blackline);
    JLabel txt = new JLabel("Press any key to Exit...");
    txt.setBounds(305,250,300,30);
    txt.setForeground(Color.RED);
    txt.setFont(new Font("Tahoma",Font.PLAIN,16));
    d2.setResizable(false);
    d2.addKeyListener(this);  
    d2.add(heading);d2.add(txt);d2.add(subheading);
    d2.add(l1);d2.add(l2);d2.add(l3);d2.add(l4);
    d2.setSize(760,320);
    d2.setLocation(330,50);
    fae.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    d2.setVisible(true);  
}
public void keyPressed(KeyEvent e) {d2.dispose();}
public void keyReleased(KeyEvent arg0) {}
public void keyTyped(KeyEvent arg0) {}
//main method
	 public static void main(String...args) throws InvalidFormatException, IOException{
		 new AboutDev();}
	
		
}
