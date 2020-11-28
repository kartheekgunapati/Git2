package Yahoo;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class yahoocode {
	
	
	public static void main(String args[]) throws Exception {
		
		Configfile config = new Configfile();
		System.setProperty("webdriver.chrome.driver", config.getChrome());
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get(config.getyahoo());
		Thread.sleep(3000);
		Pomlibrary pom = new Pomlibrary(driver);
		pom.yahooDetails();
		Thread.sleep(3000);
		driver.quit();
		
		
	}

}
