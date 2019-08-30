package pkg1;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writedata {
	
	
public static void main(String[] args) throws IOException {
	
	File f = new File("C:\\Users\\RAHUL TAAK\\Desktop\\second.xlsx");
    FileOutputStream fo = new FileOutputStream(f);
    XSSFWorkbook xw = new XSSFWorkbook();
    XSSFSheet  xs=xw.createSheet("rahul");
    
    for(int i=0;i<3;i++)
    {
    	XSSFRow xr=xs.createRow(i);
    	for(int j=0;j<6;j++)
    	{
    		XSSFCell xc = xr.createCell(j);
    		xc.setCellValue("rahul");
    	}
    }
    xw.write(fo);
    fo.flush();
	
}

}
