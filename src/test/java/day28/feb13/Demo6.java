package day28.feb13;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Demo6 {

	public static void main(String[] args)
			throws InterruptedException, EncryptedDocumentException, FileNotFoundException, IOException {
		
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.google.com/");
		driver.manage().window().maximize();
		WebElement search = driver.findElement(By.name("q"));
		search.sendKeys("lap");
		Thread.sleep(2000);
		
		WebElement list = driver.findElement(By.xpath("//div[@jsname='aajZCb']"));
		List<WebElement> rows = list.findElements(By.tagName("li"));
		int count = rows.size();
		System.out.println(count);
		
		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/GoogleSearchResult.xlsx"));
		Sheet sheet1 = wb.getSheetAt(0);
		int rownum = 0;
		for (WebElement element : rows) 
		{
			String search1 = element.getText();
			System.out.println(search1);
			sheet1.createRow(rownum++).createCell(1).setCellValue(search1);
		}
			FileOutputStream fos = new FileOutputStream("./data/GoogleSearchResult.xlsx");
			wb.write(fos);
			fos.close();
			
			driver.quit();
	}
}

