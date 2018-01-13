package POIPackage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class TestPoi
{

	@Test(dataProvider="DataInput")
	public void funtionalityTest(String userName,String pasw) throws Exception
	{
		System.setProperty("webdriver.chrome.driver",
				"E:\\DATA\\BROWSERR DRIVER\\chromedriver.exe");
	WebDriver driver=new ChromeDriver();
	driver.get("https://www.facebook.com/");
	driver.findElement(By.id("email")).sendKeys(userName);
	driver.findElement(By.id("pass")).sendKeys(pasw);
	Thread.sleep(1000);
	driver.findElement(By.xpath(".//*[@id='loginbutton']")).click();
	Thread.sleep(4000);
	driver.quit();
	
  }
	@DataProvider(name="DataInput")
	public static Iterator fetchData() throws Exception
	{
		ArrayList myData=new ArrayList();
		
			FileInputStream fis=new FileInputStream("E:\\file\\Book1.xlsx");
			Workbook wb=WorkbookFactory.create(fis);
			Sheet sh=wb.getSheet("Sheet1");
			int noOfRows=sh.getLastRowNum();
			
			String userName,pasw;
			for(int i=1;i<=noOfRows;i++)
			{
				
				userName=sh.getRow(i).getCell(0).getStringCellValue();
				pasw=sh.getRow(i).getCell(1).toString();
				
				FileOutputStream fos=new FileOutputStream("E:\\file\\Book4.xlsx");
				String result="TestPassed";
				sh.getRow(i).createCell(2).setCellValue(result);
				wb.write(fos);
				fos.close();
				
				myData.add(new Object[] {userName,pasw});
			}
				
			return myData.iterator();
				
			}
			
	}
	
	

