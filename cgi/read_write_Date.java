package cgi;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
public class read_write_Date {
	
read_write_Date() throws IOException{
 System.out.println("hey");
 Workbook wbwrite = new HSSFWorkbook();
	CreationHelper createHelper = wbwrite.getCreationHelper();
	Sheet sheet_write = wbwrite.createSheet("new sheet");
 FileInputStream myStream = new FileInputStream("demo2.xls");
 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
 HSSFSheet sheet = wb.getSheetAt(0);
 HSSFRow row;
 HSSFCell cell;
 int rowStart = sheet.getFirstRowNum() ;
    int rowEnd = sheet.getLastRowNum() ;
    int fCell,lCell;
    Row rowwrite[] =new Row[rowEnd+1];
 int tmp = 0;
  for(int i=rowStart;i<=rowEnd;i++){
	 row=sheet.getRow(i);
	 if(row==null){
    		System.out.println("empty accessed");
		continue;
		}
 if(row!=null){
	 rowwrite[i]=sheet_write.createRow((short)i);;
	 System.out.println("Last cell : " +row.getLastCellNum());
		 fCell = row.getFirstCellNum(); 
     lCell = row.getLastCellNum();
     
     for (int iCell = fCell; iCell < lCell; iCell++) {
    	 cell = row.getCell(iCell);
    	 if(cell==null){
    		 continue; 
    	 }
    	 else{
    		 Cell currentCell = cell;
    		 if (i>=0 && iCell ==0  || i>=0 && iCell ==1 ) {
    			 CellStyle style = wbwrite.createCellStyle();
    		       style.setDataFormat(
    		           createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
    		       Cell cell2 = rowwrite[i].createCell(iCell);
   	            cell2.setCellValue(row.getCell(iCell).getDateCellValue());
   	            cell2.setCellStyle(style); 
   	         sheet_write.setColumnWidth(iCell,1100*4);
        continue;
    }
    		 	else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
    		 		System.out.print(currentCell.getNumericCellValue() + "--");
    		 	}
    	 }
     }
 }
 FileOutputStream fileOut = new FileOutputStream("tryout.xls");
wbwrite.write(fileOut);
fileOut.close();
System.out.println("WorkBook has been created");
	  }
	                        
}
			
public static void main(String[] args) throws IOException{
	new read_write_Date();
}
}
