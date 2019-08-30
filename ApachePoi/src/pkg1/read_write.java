package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class read_write {
	
	public static void main(String[] args) throws IOException {
		
		File f= new File("C:\\Users\\RAHUL TAAK\\Desktop\\first.xlsx");
		FileInputStream fi = new FileInputStream(f);
		XSSFWorkbook xw=new XSSFWorkbook(fi);
		XSSFSheet xs=xw.getSheetAt(0);
		int r = xs.getPhysicalNumberOfRows();
		
		
		File f1 = new File("C:\\Users\\RAHUL TAAK\\Desktop\\third.xlsx");
		FileOutputStream fo = new FileOutputStream(f1);
		XSSFWorkbook xw1 = new XSSFWorkbook();
		XSSFSheet xs1 = xw1.createSheet("sh1");
		for(int i=0;i<r;i++)
		{
			
			XSSFRow xr=xs1.createRow(i);
			for(int j=0;j<6;j++)
			{
				XSSFCell xc = xr.createCell(j);
				xc.getCellComment();
			}
			
		}
		
		xw1.write(fo);
		fo.flush();
		
	}

}
