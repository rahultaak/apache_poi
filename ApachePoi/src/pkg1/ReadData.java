package pkg1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	
	public static void main(String[] args) throws IOException {
		
		File f= new File("C:\\Users\\RAHUL TAAK\\Desktop\\first.xlsx");
		FileInputStream fi = new FileInputStream(f);
		XSSFWorkbook fw = new XSSFWorkbook(fi);
		XSSFSheet fs = fw.getSheetAt(0);
		int r = fs.getPhysicalNumberOfRows();
		
		for(int i=0;i<r;i++)
		{
			XSSFRow fr = fs.getRow(i);
			for(int j=0;j<fr.getPhysicalNumberOfCells();j++)
			{
				XSSFCell fc = fr.getCell(j);
				System.out.println(fc.getStringCellValue());
			}
		}
		
	}

}
