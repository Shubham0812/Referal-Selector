package cgi;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class readExcel {
	readExcel() throws IOException{
		 System.out.println("hey");
		 FileInputStream myStream = new FileInputStream("demo2.xls");
		 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
		 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
		 HSSFSheet sheet = wb.getSheetAt(0);
		 HSSFRow row;
		 HSSFCell cell;
		 int rowStart = sheet.getFirstRowNum() ;
		    int rowEnd = sheet.getLastRowNum() ;
		    int fCell,lCell;
		 int tmp = 0;
		  for(int i=rowStart;i<=rowEnd;i++){
			 row=sheet.getRow(i);
			 if(row==null){
		    		System.out.println("empty accessed");
		    		continue;
		    		}
			 if(row!=null){
				 fCell = row.getFirstCellNum(); 
		         lCell = row.getLastCellNum();
		         
		         for (int iCell = fCell; iCell < lCell; iCell++) {
		        	 cell = row.getCell(iCell);
		        	 if(cell==null){
		        		 continue; 
		        	 }
		        	 else{
		        		 Cell currentCell = cell;
		        		 if (HSSFDateUtil.isCellDateFormatted(row.getCell(iCell))) {
		        		        System.out.println ("Row No.: " + row.getRowNum ()+ " " + 
		        		            row.getCell(iCell).getDateCellValue());
		        		    }
		        		 else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
		                     System.out.print(currentCell.getNumericCellValue() + "--");
		        		 }
		        	 }
		         }
			 }
		  }
		                        
	}
public static void main(String[] args) throws IOException{
	new readExcel();
}
	        	 }
			 
