package scripts;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Properties;
import java.util.Random;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContact {
	public static void main(String[] args) throws IOException, InterruptedException {
		   WebDriver driver;
		   FileInputStream fis = new FileInputStream("./src/test/resources/commondata.properties");
		   Properties p = new Properties();
		   p.load(fis);

		   String BROWSER = p.getProperty("browser");
		   if(BROWSER.equalsIgnoreCase("chrome"))
		   {
			 WebDriverManager.chromedriver().setup();
			 driver=new ChromeDriver();
		   }
		   else
		   {
			 WebDriverManager.firefoxdriver().setup();
			 driver=new FirefoxDriver();   
		   }
		   
		   String URL = p.getProperty("url");
		   String USERNAME = p.getProperty("username");
		   String PASSWORD = p.getProperty("password");
		   
		   driver.get(URL);
		   driver.findElement(By.name("user_name")).sendKeys(USERNAME);
		   driver.findElement(By.name("user_password")).sendKeys(PASSWORD);
		   driver.findElement(By.id("submitButton")).click();
		 //To avoid duplicates
		   Random ran = new Random();
		   int randname = ran.nextInt(1000);
		   driver.findElement(By.xpath("//a[text()='Contacts']")).click();
		   driver.findElement(By.xpath("//img[@alt='Create Contact...']")).click();
		   
		   FileInputStream fis1 = new FileInputStream("./src/test/resources/ExcelData.xlsx");
		   Workbook wb1 = WorkbookFactory.create(fis1);
		   Sheet sh1 = wb1.getSheet("Sheet2");
		   Row row1 = sh1.getRow(0);
		   Cell cel1 = row1.getCell(0);
		   String contname = cel1.getStringCellValue()+randname;
		   driver.findElement(By.name("lastname")).sendKeys(contname);
		   driver.findElement(By.name("button")).click();
		   driver.findElement(By.xpath("(//td[@class='small'])[2]")).click();
	       driver.findElement(By.linkText("Sign Out")).click();
}
}