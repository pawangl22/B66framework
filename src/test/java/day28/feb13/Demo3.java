package day28.feb13;
import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo3 {

	public static void main(String[] args) throws Exception 
	{

		Workbook wb = WorkbookFactory.create(new File("./data/Selenium.xlsx"));
		for(int i=0;i<5;i++)
		{
			for(int j=0;j<3;j++) 
			{
				String v = wb.getSheet("sheet2").getRow(i).getCell(j).getStringCellValue();
				System.out.print(v+" ");
			}
			System.out.println();
		}
		wb.close();
		
	}

}
