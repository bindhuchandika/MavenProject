package scripts;

import java.io.FileInputStream;
import java.util.Properties;
import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateOrg {

	public static void main(String[] args) throws Throwable {
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
    driver.manage().window().maximize();
	String URL = p.getProperty("url");
	String USERNAME = p.getProperty("username");
	String PASSWORD = p.getProperty("password");
    driver.get(URL);
	driver.findElement(By.name("user_name")).sendKeys(USERNAME);
	driver.findElement(By.name("user_password")).sendKeys(PASSWORD);
	driver.findElement(By.id("submitButton")).click();

	Random ran = new Random();
	int randname = ran.nextInt(1000);
	driver.findElement(By.linkText("Organizations")).click();
	driver.findElement(By.xpath("//img[@alt='Create Organization...']")).click();
	
	FileInputStream fis1 = new FileInputStream("./src/test/resources/ExcelData.xlsx");
	Workbook wb1 = WorkbookFactory.create(fis1);
    Sheet sh1 = wb1.getSheet("Sheet1");
    Row row1 = sh1.getRow(0);
    Cell cel1 = row1.getCell(0);
    String orgname = cel1.getStringCellValue()+randname;
    driver.findElement(By.name("accountname")).sendKeys(orgname);
	driver.findElement(By.name("button")).click();
	driver.navigate().refresh();
	driver.findElement(By.xpath("(//td[@class='small'])[2]")).click();
    driver.findElement(By.linkText("Sign Out")).click();

	   
	}
}