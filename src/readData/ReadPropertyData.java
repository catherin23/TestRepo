package readData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadPropertyData {

	public static void main(String[] args) throws Exception {
		
		File f = new File("C:\\Users\\Hp\\eclipse-workspace\\DataDrivenFramework\\env.properties");
		FileInputStream fis = new FileInputStream(f);
		
		Properties prop = new Properties();
		prop.load(fis);
		
		
		
		System.setProperty("webdriver.chrome.driver", "E:/Drivers/chromedriver.exe");
		WebDriver driver= new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(prop.getProperty("url"));
		
		
		WebElement search = driver.findElement(By.name("q"));
		search.sendKeys(prop.getProperty("searchterm"));
		search.sendKeys(Keys.ENTER);
		
		driver.quit();
	

	}

}
