package cgi;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.beans.PropertyChangeEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.SwingConstants;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.border.Border;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Intro{
JFrame f;
JLabel ta,label1,label2,label3,label4,label5,error;
JButton b,ba,b1,b2,b2x,b3,b6,submit,Vb1,Vb2,b4,b5,Vb3,Vb4,Vb5,back,executeAll;
JProgressBar jb;
JProgressBar progress;
private static JDialog d,d1,d2;
String result,result2;
String outputFile1,outputFile2,inputFile3,inputFile4;
String in1,in2,in3;
String toCheckMaster,toCheckCandidate,toCheckAMS;
int i=0,num=0,count=0,count2=0,numbercounter = 7,count3 =0;    
int code,cou = 0;
//constructor
Intro() throws IOException, InvalidFormatException {

		JFrame.setDefaultLookAndFeelDecorated(true);	
		f = new JFrame("Master Referral Validation Automator");
		f.setLayout(null);
		b = new JButton("Format Master Tracker: Select Master Tracker File");
		b.setBounds(460, 130, 350, 30);
		b.setMnemonic(KeyEvent.VK_M);
		ba = new JButton("Format Candidate Referrals : Candidate File");
		ba.setBounds(460, 130, 350, 30);
		ba.setVisible(false);
		ba.setMnemonic(KeyEvent.VK_C);
		b.setVisible(false);
		b1 = new JButton("Select Master Tracker File: ");
		b1.setBounds(460, 130, 250,30);
		b1.setVisible(false);
		b1.setMnemonic(KeyEvent.VK_M);
		b2 = new JButton("Select Candidate Referral File: ");
		b2.setBounds(460, 200, 250,30);
		b2.setVisible(false);
		b2.setMnemonic(KeyEvent.VK_C);
		b2x = new JButton("Select AMS Dump File: ");
		b2x.setBounds(460, 270, 250,30);
		b2x.setVisible(false);
		b2x.setMnemonic(KeyEvent.VK_A);
		b3 = new JButton("Module 1: ");
		b3.setBounds(10, 130, 150, 30);
		b3.setMnemonic(KeyEvent.VK_1);
		b3.setToolTipText("The output file contains the MR Requisition File along with Referred By Name and Email from Candidate Referrals");
		b4 = new JButton("Module 2: ");
		b4.setBounds(10, 200, 150, 30);
		b4.setMnemonic(KeyEvent.VK_2);
		b4.setToolTipText("The Output file contains Duplicacy Check along with ID, Source, Current Stage & Current Status");
		b5 = new JButton("Module 3: ");
		b5.setBounds(10, 270, 150, 30);
		b5.setMnemonic(KeyEvent.VK_3);
		b5.setToolTipText("The Output file contains Communication Mails for different Sources, Stage & Status ");
		executeAll = new JButton("ExecuteAll");
		executeAll.setMnemonic(KeyEvent.VK_E);
		executeAll.setBounds(820, 320, 150, 30);
		Color color = UIManager.getColor("f.background");
		back = new JButton("Back",new ImageIcon("C:\\Users\\shubham.k.singh\\Desktop\\cgi\\cgi\\back.png"));
		back.setBounds(20, 20, 130, 30);
		back.setBackground(color);
		back.setMnemonic(KeyEvent.VK_B);
		back.setVisible(false);
		submit = new JButton("Submit Selected File(s)");
		submit.setBounds(370,390,200,30);
		submit.setVisible(false);
		submit.setMnemonic(KeyEvent.VK_S);
		Vb1= new JButton("Select Master Tracker File");
		Vb1.setBounds(460,130,210,30);
		Vb1.setVisible(false);
		Vb1.setMnemonic(KeyEvent.VK_M);
		Vb2= new JButton("Select Candiate Referral File");
		Vb2.setBounds(680,130,210,30);
		Vb2.setVisible(false);
		Vb2.setMnemonic(KeyEvent.VK_C);
		Vb3= new JButton("Select MR OutputFile");
		Vb3.setBounds(460,200,210,30);
		Vb3.setVisible(false);
		Vb3.setMnemonic(KeyEvent.VK_M);
		Vb4= new JButton("Select AMS Dump File");
		Vb4.setBounds(680,200,210,30);
		Vb4.setVisible(false);
		Vb4.setMnemonic(KeyEvent.VK_E);
		Vb5= new JButton("Select MR & AMS Output File");
		Vb5.setBounds(680,270,210,30);
		Vb5.setVisible(false);
		Vb5.setMnemonic(KeyEvent.VK_F);
		label1 = new JLabel("Format Master Tracker");
		//label1.setBounds(190, 130, 500, 30);
		//label1.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label2 = new JLabel("Format Candidate Referral");
		//label2.setBounds(190, 200, 500, 30);
		//label2.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label3 = new JLabel("Perform VLookup");
		label3.setBounds(190, 130, 500, 30);
		label3.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label4 = new JLabel("VLookup From AMS Dump");
		label4.setBounds(190, 200, 500, 30);
		label4.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		label5 = new JLabel("Get Communication Mails");
		label5.setBounds(190, 270, 500, 30);
		label5.setFont(new Font("Times New Roman",Font.LAYOUT_LEFT_TO_RIGHT, 18));
		error = new JLabel("",SwingConstants.CENTER);
		error.setBounds(10, 430, 950, 30);
		error.setFont(new Font("Courier New", Font.BOLD, 20));
		ta = new JLabel("Member Referral Validation Automator",SwingConstants.CENTER);
		ta.setBounds(215,10,600,80); 
		Border blackline = BorderFactory.createLineBorder(Color.black);
		ta.setBorder(blackline);
		ta.setFont(new Font("Tahoma", Font.BOLD, 26));
		progress = new JProgressBar(0);
		progress.setBounds(0,350,1000,20);
		progress.setValue(0);
		progress.setStringPainted(true);
		progress.setVisible(false); 
		JMenuBar menuBar = new JMenuBar();
		JMenu optionsMenu = new JMenu("Options");
		optionsMenu.setMnemonic('O');
		JMenu helpMenu = new JMenu("Help");
		helpMenu.setMnemonic('H');
		JMenu aboutMenu = new JMenu("About");
		aboutMenu.setMnemonic(KeyEvent.VK_A);
		JMenuItem fmt = new JMenuItem("Format Master Tracker");
		fmt.setMnemonic(KeyEvent.VK_M);
		fmt.setActionCommand("format master tracker");
		optionsMenu.add(fmt);
	
		JMenuItem fcr = new JMenuItem("Format Candidate Referrals");
		fcr.setMnemonic(KeyEvent.VK_C);
		optionsMenu.add(fcr);
		
		JMenuItem helpHowTo = new JMenuItem("How to use the Application");
		helpHowTo.setMnemonic(KeyEvent.VK_H);
		helpMenu.add(helpHowTo);
		
		JMenuItem helpMT = new JMenuItem("Check Master Tracker Format");
		helpMT.setMnemonic(KeyEvent.VK_M);
		helpMenu.add(helpMT);
		
		JMenuItem helpCR = new JMenuItem("Check Candidate Referral Format");
		helpCR.setMnemonic(KeyEvent.VK_C);
		helpMenu.add(helpCR);
		
		JMenuItem helpAMS = new JMenuItem("Check AMS Dump Format");
		helpAMS.setMnemonic(KeyEvent.VK_A);
		helpMenu.add(helpAMS);
		
		JMenuItem about = new JMenuItem("About Application");
		about.setMnemonic(KeyEvent.VK_N);
		aboutMenu.add(about);
		
		menuBar.add(optionsMenu);
		menuBar.add(helpMenu);
		menuBar.add(aboutMenu);
		f.setJMenuBar(menuBar);
		f.setResizable(false);
		f.add(b);f.add(label1);f.add(ta);f.add(b1);f.add(b2);f.add(b3);f.add(ba);f.add(label2);f.add(label3);f.add(error);f.add(Vb1);f.add(Vb2);
		f.add(submit);f.add(b4);f.add(b5);f.add(label4);f.add(label5);f.add(Vb3);f.add(Vb4);f.add(Vb5);f.add(executeAll);
		f.setSize(1000,520);f.add(back);f.add(progress);f.add(b2x);
		f.getContentPane().setBackground(new Color(255,255,255));
		f.setLocation(200,20);
		f.setVisible(true);
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		back.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        finish();}			
	     });

		about.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        new AboutDev();}			
	     });
		
		helpHowTo.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	new helper();}			
	     });
		helpMT.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	helpMasterController();}			
	     });
		
		helpCR.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	helpCandidateController();}			
	     });
		
		helpAMS.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	helpAMSController();}			
	     });
		
		executeAll.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	b3.setEnabled(false);
	        	b4.setEnabled(false);
	        	b5.setEnabled(false);	
	        	b1.setVisible(true);
	        	back.setVisible(true);
	        	b2.setVisible(true);
	        	b2x.setVisible(true);
	        	//submit.setVisible(true);
	        }			
	     });
		
		b1.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {

	        	//provide user to select the file
	        	 in1 = selectfile();
	        	 if(in1==null){
	        	 error.setText("Master Tracker File Not Chosen");
	        	 
	        	 return;
	        	 }
	        	 count3+=1;
	        	 	if(count3==2){
	        	 		code = 6;
	        	 		submit.setVisible(true);
	        	 	}
	        	 	error.setText("Master Tracker File Selected");

	        	
	        	}			
	     });
		
		b2.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {

	        	//provide user to select the file
	        	 in2 = selectfile();
	        	 if(in2==null){
	        	 error.setText("Candidate Referrals File Not Chosen");
	        	 return;
	        	 }
	        	 count3+=1;
	        	 	if(count3==3){
	        	 		code = 6;
	        	 		submit.setVisible(true);
	        	 	}
	        	 	error.setText("Candidate Referrals File Selected");

	        	
	        	}			
	     });
		
		b2x.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {

	        	//provide user to select the file
	        	 in3 = selectfile();
	        	 if(in3==null){
	        	 error.setText("AMS Dump File Not Chosen");
	        	 return;
	        	 }
	        	 	count3+=1;
	        	 	if(count3==3){
	        	 		code = 6;
	        	 		submit.setVisible(true);
	        	 	}
	        	 	error.setText("AMS Dump File Selected");

	        	
	        }			
	     });
		fmt.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        b.setVisible(true);
	        error.setText("");
	        back.setVisible(true);
	        b2.setEnabled(false);
	        b3.setEnabled(false);
	        b4.setEnabled(false);
	        b5.setEnabled(false);}			
	     });
		fcr.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        error.setText("");
	        ba.setVisible(true);
	        back.setVisible(true);
	        b1.setEnabled(false);
 	        b3.setEnabled(false);
	        b4.setEnabled(false);
	        b5.setEnabled(false);}			
	     });
		
		b3.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        try {
	        	error.setText("");
		        back.setVisible(true);
		        executeAll.setEnabled(false);
		        b1.setEnabled(false);
		        b2.setEnabled(false);
		        b4.setEnabled(false);
		        b5.setEnabled(false);
		        Vb1.setVisible(true);
		        Vb2.setVisible(true);
				//new SheetCopy();
				//finish();
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}}			
	     });
		b4.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        try {
	        	error.setText("");
		        back.setVisible(true);
		        executeAll.setEnabled(false);
		        b1.setEnabled(false);
		        b2.setEnabled(false);
		        b3.setEnabled(false);
		        b5.setEnabled(false);
		        Vb3.setVisible(true);
		        Vb4.setVisible(true);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}}			
	     });
		b5.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        error.setText("");
	        executeAll.setEnabled(false);
	        back.setVisible(true);
	        Vb5.setVisible(true);
	        b1.setEnabled(false);
	        b2.setEnabled(false);
	        b3.setEnabled(false);
	        b4.setEnabled(false);}			
	     });
		
		
		Vb1.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        	  outputFile1 = selectfile();
	        	 if(outputFile1==null){
	        	 error.setText("Master Tracker File Not Chosen");
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
	        	 error.setText("Candidate Referral File Not Chosen");
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
		Vb3.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        	  inputFile3 = selectfile();
	        	 if(inputFile3==null){
	        	 error.setText("MR Output File Not Chosen");
	        	 count2+=1;
	        	 return;
	        	 }
	        	 	error.setText("MR Output File Selected");
	        }
	     });
		Vb4.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        				inputFile4 = selectfile();
	        	 if(inputFile4==null){
	        	 error.setText("AMS Dump File Not Chosen");
	        	 count2+=1;
	        	 return;
	        	 }
	        	 error.setText("AMS Dump File Selected");
	        	 if(inputFile3!=null&&inputFile4!=null){
	        		   code=4;
	        		 submit.setVisible(true);
	        	        	 }
					 }
	     });
		Vb5.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        				error.setText("");
	        	//provide user to select the file
	        	  result = selectfile();
	        	 if(result==null){
	        	 error.setText("No File Chosen");
	        	 finish();
	        	 return;
	        	 }
	        		 code=5;
	        		 submit.setVisible(true);
					 }
	     });
		b.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
        				error.setText("");
        	//provide user to select the file
        				
        	  result = selectfile();
        	 if(result==null){
        	 error.setText("No File Chosen");
        	 finish();
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
	            	 error.setText("No File Chosen");
	        	 finish();
	        	 return;
	        	 }
	        		 code=2;
	        		 submit.setVisible(true);

					 }
	     });
		
		submit.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	
	        	if(code==1){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){
	        			public void run() {
	        				b1.setEnabled(false);
	        				submit.setEnabled(false);
	        				executeAll.setEnabled(false);
	        				error.setText("Please Wait");
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(300);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
//	        							finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	      	        		read_write(result);			
	      					
	      					
	        		  } catch (IOException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
	        			   }
	        	
	        	if(code==2){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){
	        			public void run() {
	        				b2.setEnabled(false);
	        				submit.setEnabled(false);
	        				executeAll.setEnabled(false);
	        				error.setText("Please Wait");
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(700);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
	        						//	finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	        				  candidate_referrals obj = new candidate_referrals();
	        				  obj.modify(result2);
	      					
	        		  } catch (IOException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
	        	}
	        	if(code==3){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){
	        			public void run() {
	        				b3.setEnabled(false);
	        				Vb1.setEnabled(false);
	        				Vb2.setEnabled(false);
	        				submit.setEnabled(false);
	        				error.setText("Please Wait");
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(1100);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
	        							//finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	        				  new SheetCopy(outputFile1,outputFile2);	
	        				
	      					
	        		  } catch (IOException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
	        	}
	        	if(code==4){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){
	        			public void run() {
	        				b4.setEnabled(false);
	        				Vb4.setEnabled(false);
	        				Vb3.setEnabled(false);
	        				submit.setEnabled(false);
	        				error.setText("Please Wait");
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(1300);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
	        				//			finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	        				  new AMSdump(inputFile3,inputFile4); 
	        				
	      					
	        		  } catch (IOException | InvalidFormatException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
	      }
	        	if(code==5){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){
	        			public void run() {
	        				b5.setEnabled(false);
	        				Vb5.setEnabled(false);
	        				
	        				submit.setEnabled(false);
	        				error.setText("Please Wait");
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(20);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
	        							//finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	        				  new Formatting(result);
	        				
	      					
	        		  } catch (IOException | InvalidFormatException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
			
	        			   }
	        	
				
	    		if(code==6){
	        		progress.setVisible(true);
	        		Runnable run = new Runnable(){

	        			public void run() {
		        			b1.setEnabled(false);
		        			b2.setEnabled(false);
		        			b2x.setEnabled(false);
	        				 for(int i =0;i<101;i++){
	        					 progress.setValue(i);
	        					 try {
	        						Thread.sleep(2940);
	        						if(i==100){
	        							error.setText("The Work has been Finished");
	        							//finish();

	        						}
	        					} catch (InterruptedException e) {
	        						// TODO Auto-generated catch block
	        						e.printStackTrace();
	        					}   
	        				 }

	        			}
	        		};
	        		Thread  t = new Thread(run);
	        		t.start();
	        	//	t[cou].start();
	        	  Runnable rua = new Runnable(){
	        		  public void run(){
	        			  try{
	        				  back.setVisible(false); 
	        				  error.setText("Please Wait While The Applicaiton is Working...");
	  	    					checkExecutePress(in1,in2,in3);	 	
	      					
	        		  } catch (IOException | InvalidFormatException e1) {}
	        		  }  
	        	  };
	        	    Thread t1 = new Thread(rua);
	        	t1.start();
	    		}

	      }
	     });

		
}
	

public void helpMasterController(){
	JFrame fa= new JFrame();  
	
	  d = new JDialog(fa , "Master Tracker Format Verify", true);  
      d.setLayout(null);
      JLabel label = new JLabel("Choose The Master Tracker Input File");
      label.setFont(new Font("Agency FB",Font.BOLD,24));
      label.setBounds(10,10,300,30);
      JButton b = new JButton ("Select File");
      b.setBounds(330,30,150,30);
      b.setMnemonic(KeyEvent.VK_F);
      JButton sub = new JButton ("Submit");
      sub.setBounds(240,130,150,30);
      sub.setMnemonic(KeyEvent.VK_S);
      sub.setVisible(false);
      b.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
              toCheckMaster = selectfile();
              if(toCheckMaster==null){
            	  JOptionPane.showMessageDialog(null, "File not Selected");
              }
              else{
            	  sub.setVisible(true);
            	 
              }
          }  
      });  
      sub.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
        	 int res = helpMaster.checkFormat(toCheckMaster);
        	 if(res==19){
        		 JOptionPane.showMessageDialog(null, "Yes, The Format of Master Tracker is Correct, You Can Continue");
        		 d.dispose();
        		 toCheckMaster="";
        	 }
        	 else if(res==-1){
        		 d.dispose();
        		 toCheckMaster="";
        	 }
        	 else{
        		 toCheckMaster="";
        		 d.dispose();
        		 new helpMaster();
        	 }
        		 
          }  
      });  
      d.add( new JLabel ("Click button to continue."));  
      d.add(b);d.add(label);d.add(sub);
      d.setSize(500,200);
      d.setLocation(230,80);
      d.setVisible(true);  
}


public void helpCandidateController(){
	JFrame faa= new JFrame();  
	
	  d1 = new JDialog(faa , "Candidate Referral Format Verify", true);  
      d1.setLayout(null);
      JLabel label = new JLabel("Choose The Candidate Referral Input File");
      label.setFont(new Font("Agency FB",Font.BOLD,24));
      label.setBounds(10,10,350,30);
      JButton b = new JButton ("Select File");
      b.setBounds(346,40,150,30);
      b.setMnemonic(KeyEvent.VK_F);
      JButton suba = new JButton ("Submit");
      suba.setBounds(240,130,150,30);
      suba.setMnemonic(KeyEvent.VK_S);
      suba.setVisible(false);
      b.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
              toCheckCandidate = selectfile();
              if(toCheckCandidate==null){
            	  JOptionPane.showMessageDialog(null, "File not Selected");
              }
              else{
            	  suba.setVisible(true);
            	 
              }
          }  
      });  
      suba.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
        	 int res = helpCandidate.checkFormat(toCheckCandidate);
        	 if(res==24){
        		 JOptionPane.showMessageDialog(null, "Yes, The Format of Candidate Referral is Correct, You Can Continue");
        		 d1.dispose();
        		 toCheckCandidate="";
        	 }
        	 else if(res==-1){
        		 toCheckCandidate="";
        		 d1.dispose();
        	 }
        	 else{
        		 toCheckCandidate="";
        		 d1.dispose();
        		 new helpCandidate();
        	 }
        		 
          }  
      });  
      d1.add( new JLabel ("Click button to continue."));  
      d1.add(b);d1.add(label);d1.add(suba);
      d1.setSize(500,200);
      d1.setLocation(230,80);
      d1.setVisible(true);  
}

public void helpAMSController(){
	JFrame faae= new JFrame();  
	
	  d2 = new JDialog(faae , "AMS Dump Format Verify", true);  
      d2.setLayout(null);
      JLabel label = new JLabel("Choose The Ams Dump File");
      label.setFont(new Font("Agency FB",Font.BOLD,24));
      label.setBounds(10,10,350,30);
      JButton b = new JButton ("Select File");
      b.setBounds(330,40,150,30);
      b.setMnemonic(KeyEvent.VK_F);
      JButton suba2 = new JButton ("Submit");
      suba2.setBounds(240,130,150,30);
      suba2.setMnemonic(KeyEvent.VK_S);
      suba2.setVisible(false);
      b.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
              toCheckAMS = selectfile();
              if(toCheckAMS==null){
            	  JOptionPane.showMessageDialog(null, "File not Selected");
              }
              else{
            	  suba2.setVisible(true);
            	 
              }
          }  
      });  
      suba2.addActionListener ( new ActionListener()  
      {  
          public void actionPerformed( ActionEvent e )  
          {  
			int res = helpAms.checkFormat(toCheckAMS);
        	 if(res==7){
        		 JOptionPane.showMessageDialog(null, "Yes, The Format of AMS Dump is Correct, You Can Continue");
        		 d2.dispose();
        		 toCheckAMS="";
        	 }
        	 else if(res==-1){
        		 toCheckAMS="";
        		 d2.dispose();
        	 }
        	 else{
        		 toCheckAMS="";
        		 System.out.println(res);
        		 d2.dispose();
        		 new helpAms();
        	 }
        		 
          }  
      });  
      d2.add(b);d2.add(label);d2.add(suba2);
      d2.setSize(500,200);
      d2.setLocation(230,80);
      d2.setVisible(true);  
}
//to check the contents of master tracker file





//
public void checkExecutePress(String Master,String Candidate,String AmsDump) throws IOException, InvalidFormatException{
	

	
	back.setEnabled(false);
	executeAll.setEnabled(false);
	new SheetCopy(Master,Candidate);
	error.setText("VlookUp From Master Tracker and Candidate Referral Done");
	new AMSdump("VLookupOutputs.xlsx",AmsDump);
	error.setText("Duplicacy Check along with ID, Source, Current Stage & Current Status Done");
	new Formatting("AmsDumpOutput.xlsx");
	//error.setText(" Communication Mails for different Sources, Stage & Status Done");
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
	Workbook wbwrite = new XSSFWorkbook();
	CreationHelper createHelper = wbwrite.getCreationHelper();
	
	Sheet sheet_write = wbwrite.createSheet("Sheet1");
	wbwrite.createSheet("Sheet2");
	wbwrite.createSheet("Sheet3");
	FormulaEvaluator evaluator = wbwrite.getCreationHelper().createFormulaEvaluator();
	
	CellStyle num = wbwrite.createCellStyle();
		num.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
	
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
    String number_c;
	//font style to set font as bold
	//code to iterate over the rows  
    
	for(int i=rowStart;i<=rowEnd;i++){
	row=sheet.getRow(i);
	if(row==null){
		//System.out.println("empty accessed");
		continue;
	}
	if(row!=null){
		rowwrite[i]=sheet_write.createRow((short)i);
		//first and last cell for the row
		 fCell = row.getFirstCellNum(); 
         lCell = row.getLastCellNum();	////System.out.println("First :  " + fCell + "Last : " + lCell);
         for(int iCell = fCell; iCell < lCell; iCell++) {
         cell = row.getCell(iCell);
		 if(cell==null){
			 continue;
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
    			 if(i>=6&& iCell ==9){
    				 double value = currentCell.getNumericCellValue();
    				 String axe =""+currentCell.getAddress();
    				 if(axe.length()==2){
    				 number_c = axe.substring(1,2);
    				 }else if(axe.length()==3){
    				 number_c = axe.substring(1,3);
    				 }
    				 else{
    					 number_c=axe.substring(1,4);
    				 }
    				 //System.out.println("hehe  " + axe + number_c);
    				 rowwrite[i] = sheet_write.getRow((short)i);
    				 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+value+",10)");
    				 
    				 CellReference cellReference = new CellReference("K"+number_c);
    				 Row rowF = sheet_write.getRow(cellReference.getRow());
    	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
    	         		//System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
    	         		CellValue cellValue = evaluator.evaluate(cellF);
  
    	         	Cell xcu =rowwrite[i].createCell(iCell+1);
    	         	xcu.setCellStyle(num);
    	         //	long final_result = Integer.parseInt(cellValue.getStringValue());
    	    //     	//System.out.println(final_result);
	         		//System.out.println("  "+cellValue.getStringValue());
    	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
    				 continue;
    				 
    			 }
                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getNumericCellValue());    
				 }
			 
			 else if (currentCell.getCellTypeEnum() == CellType.STRING) {
		//		 //System.out.print(currentCell.getStringCellValue() + "--");
    			 if(i>=6&& iCell ==9){
    				 
    				 try{
    					 Row are = sheet.getRow(i);
    					 //System.out.println("huh" + are.getCell(9).getStringCellValue()+"a");
    					 if(are.getCell(9).getStringCellValue().equals(" ")){
    						 {
    							 try{
    								 Cell currentCells = row.getCell(12);
    								 if (currentCells.getCellTypeEnum() == CellType.NUMERIC) {
    								 double value = currentCells.getNumericCellValue();
    			    				 String axe =""+currentCells.getAddress();
    			    				 if(axe.length()==2){
    			    				 number_c = axe.substring(1,2);
    			    				 }else if(axe.length()==3){
    			    				 number_c = axe.substring(1,3);
    			    				 }
    			    				 else{
    			    					 number_c=axe.substring(1,4);
    			    				 }
    			    				 rowwrite[i] = sheet_write.getRow((short)i);
    			    				 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+value+",10)");
    			    				 
    			    				 CellReference cellReference = new CellReference("K"+number_c);
    			    				 Row rowF = sheet_write.getRow(cellReference.getRow());
    			    	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
    			    	         		//System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
    			    	         		CellValue cellValue = evaluator.evaluate(cellF);
    			  
    			    	         	Cell xcu =rowwrite[i].createCell(iCell+1);
    			    	         	xcu.setCellStyle(num);
    			    	         //	long final_result = Integer.parseInt(cellValue.getStringValue());
    			    	    //     	//System.out.println(final_result);
    				         		//System.out.println("  "+cellValue.getStringValue());
    			    	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
    			    				 continue;
    								 }
    								 else if (currentCells.getCellTypeEnum() == CellType.STRING) {
    				    				 String add =""+currentCell.getAddress();
    				    				 if(add.length()==2){number_c = add.substring(1,2);}
    				    				 else if(add.length()==3){number_c = add.substring(1,3);}
    				        		     else{number_c=add.substring(1,4);}
    				    				 String value = currentCells.getStringCellValue();
    				    				 try{
    				    					 String newValue = value.replaceAll("-","");
    				    					 rowwrite[i] = sheet_write.getRow((short)i);
    				    					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
    				    					 CellReference cellReference = new CellReference("K"+number_c);
    				        				 Row rowF = sheet_write.getRow(cellReference.getRow());
    				        	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
    				        	         		CellValue cellValue = evaluator.evaluate(cellF);
    				        	         		//System.out.println("  "+cellValue.getStringValue());
    				        	               	Cell xcu =rowwrite[i].createCell(iCell+1);
    				            	         	xcu.setCellStyle(num);
    				            	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
    				            	         	continue;
    				    					 }catch(Exception e){
    				    					 String newValue = value.replaceAll("\\s","");
    				    					 try{
    				    						 rowwrite[i] = sheet_write.getRow((short)i);
    				        					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
    				        					 CellReference cellReference = new CellReference("K"+number_c);
    				            				 Row rowF = sheet_write.getRow(cellReference.getRow());
    				            	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
    				            	         		CellValue cellValue = evaluator.evaluate(cellF);
    				            	         		//System.out.println("  "+cellValue.getStringValue());
    				            	               	Cell xcu =rowwrite[i].createCell(iCell+1);
    				                	         	xcu.setCellStyle(num);
    				                	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
    				        				 }catch(Exception af){}
    				    					 }
    				    					 continue;
    								 }
    								 
    							 }catch(NullPointerException nula){} 
    						 }
    					 }
    				 }catch(NullPointerException a){
    					 //System.out.println("I value = " + i + "haha");
    				 }
    				 
    				 
    				 String add =""+currentCell.getAddress();
    				 if(add.length()==2){number_c = add.substring(1,2);}
    				 else if(add.length()==3){number_c = add.substring(1,3);}
        		     else{number_c=add.substring(1,4);}
    				 String value = currentCell.getStringCellValue();
    				 try{
    					 String newValue = value.replaceAll("-","");
    					 rowwrite[i] = sheet_write.getRow((short)i);
    					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
    					 CellReference cellReference = new CellReference("K"+number_c);
        				 Row rowF = sheet_write.getRow(cellReference.getRow());
        	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
        	         		CellValue cellValue = evaluator.evaluate(cellF);
        	         		//System.out.println("  "+cellValue.getStringValue());
        	               	Cell xcu =rowwrite[i].createCell(iCell+1);
            	         	xcu.setCellStyle(num);
            	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
            	         	continue;
    					 }catch(Exception e){
    					 String newValue = value.replaceAll("\\s","");
    					 try{
    						 rowwrite[i] = sheet_write.getRow((short)i);
        					 rowwrite[i].createCell(9+1).setCellFormula("RIGHT("+newValue+",10)");
        					 CellReference cellReference = new CellReference("K"+number_c);
            				 Row rowF = sheet_write.getRow(cellReference.getRow());
            	         		Cell cellF = rowF.getCell(cellReference.getCol()); 
            	         		CellValue cellValue = evaluator.evaluate(cellF);
            	         		//System.out.println("  "+cellValue.getStringValue());
            	               	Cell xcu =rowwrite[i].createCell(iCell+1);
                	         	xcu.setCellStyle(num);
                	         	xcu.setCellValue(Double.parseDouble(cellValue.getStringValue()));
        				 }catch(Exception af){}
    					 }
    					 continue;}
                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getStringCellValue());}
				 else if(currentCell.getCellTypeEnum() == CellType.FORMULA){
					 System.out.print(currentCell.getStringCellValue() + "--");
                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getCellFormula());}
				 else if (currentCell.getCellTypeEnum() == CellType.ERROR){
                  System.out.print(currentCell.getStringCellValue() + "--");
                     rowwrite[i].createCell(iCell+1).setCellValue(currentCell.getErrorCellValue());}
			 
			 }}//cell for loop ends
  
         //Validation Index Calculation
         	if(i>=6){
         		rowwrite[i]=sheet_write.getRow((short)i);;
         		rowwrite[i].createCell(0).setCellFormula("CONCATENATE(F"+counter1+",D"+counter2+")");
         		
         		CellReference cellReference = new CellReference("A"+counter1);
         		Row rowF = sheet_write.getRow(cellReference.getRow());
         		Cell cellF = rowF.getCell(cellReference.getCol()); 
         		//System.out.print(cellReference.getRow() + "  " + cellReference.getCol());
         		CellValue cellValue = evaluator.evaluate(cellF);
         		//System.out.println("  "+cellValue.getStringValue());
         		rowwrite[i].createCell(0).setCellValue(cellValue.getStringValue());
         		counter1+=1;counter2+=1;}
         	if(i==5){
         		rowwrite[i]=sheet_write.getRow((short)i);;
         		rowwrite[i].createCell(0).setCellValue("Validation Index");}}//row not null ends
	
	 System.out.println(result);
      System.out.println("WorkBook has been created");
      }//row ends
	  String path = result.replaceAll(".xls","");
	  FileOutputStream fileOut = new FileOutputStream(path+"(Output1).xlsx");
      wbwrite.write(fileOut);
      fileOut.close();
	  wbwrite.close();
	  wb.close();
	  fs.close();
	}catch(Exception e)
	{
		error.setText(e+"Invalid File Selected");
		submit.setVisible(false);
		b.setVisible(false);
   	 	b2.setEnabled(true);
   	 	b3.setEnabled(true);
   	 	b4.setEnabled(true);
   	 	b5.setEnabled(true);
		return;
		}
}

public void finish(){
	b1.setEnabled(true);
	b1.setVisible(false);
	b2.setEnabled(true);
	b2.setVisible(false);
	b2x.setVisible(false);
	b3.setEnabled(true);
	b4.setEnabled(true);
	b5.setEnabled(true);
	submit.setVisible(false);
	back.setVisible(false);
	Vb1.setVisible(false);
	Vb2.setVisible(false);
	Vb3.setVisible(false);
	Vb4.setVisible(false);
	Vb5.setVisible(false);
	progress.setValue(0);
	progress.setVisible(false);
	b.setVisible(false);
	executeAll.setEnabled(true);
	ba.setVisible(false);
	
}



public static void main(String[] args) throws IOException, InvalidFormatException{
new Intro();
}




}
