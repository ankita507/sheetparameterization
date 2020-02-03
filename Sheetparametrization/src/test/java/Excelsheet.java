//import java.util.concurrent.TimeUnit;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Excelsheet {
    WebDriver driver;
	@BeforeTest
	public void beforeTest()
	{
		System.setProperty("webdriver.chrome.driver", "C:\\selenium_drivers\\chromedriver\\chromedriver.exe");
		driver=new ChromeDriver();  //window will open in mimimise mode
		
		driver.get("http://demowebshop.tricentis.com/");
		
		driver.manage().window().maximize();    //maximise the window
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	}
	@Test
	public void testAut() 
	{
		readWriteExcel();
	}
	
	
	@AfterTest
	public void afterTest()
	{
		driver.close();
	}
	
	
	
	public void readWriteExcel() {
		File file = new File("Testdata.xlsx");
	
			InputStream is;
			XSSFWorkbook wb;
			try {
				is = new FileInputStream(file);
				wb = new XSSFWorkbook(is);
			
			XSSFSheet sheet1 = wb.getSheet("Sheet1");
			
			for(int i=1;i<=sheet1.getLastRowNum();i++)
			{
				String Username = sheet1.getRow(i).getCell(0).getStringCellValue();
				String Password = sheet1.getRow(i).getCell(1).getStringCellValue();
				String Result = login(Username, Password);
				sheet1.getRow(i).createCell(2).setCellValue(Result);
			}
			is.close();
			OutputStream os= new FileOutputStream(file);
			wb.write(os);
			wb.close();
			os.close();
			}
			catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	
	
	}
	public String login(String Username, String Password) {
		driver.findElement(By.linkText("Log in")).click();
		driver.findElement(By.name("Email")).sendKeys(Username);
		driver.findElement(By.name("Password")).sendKeys(Password);
		driver.findElement(By.xpath("//input[@value='Log in']")).click();
		
		if (driver.findElements(By.id("vote-poll-1")).size()>0) {
			String uname = driver.findElement(By.xpath("//a[@href='/customer/info']")).getText();
			if(uname.contentEquals(Username))
				driver.findElement(By.xpath("//a[@href='/logout']")).click();
			
			
		}
		else
		{
			driver.findElement(By.xpath("//a[@href='/login']")).click();
			return "Invalid user";
			
		}
		return "Valid user";
		}
	
	
	
	
}
