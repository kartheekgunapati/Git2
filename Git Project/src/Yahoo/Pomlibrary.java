package Yahoo;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;



public class Pomlibrary {
	
	WebDriver driver;
	
	public Pomlibrary(WebDriver driver) {
		this.driver = driver;
	}
	
	By signin = By.xpath("//span[text()='Sign in']");
	By createacct = By.id("createacc");
	By firstname = By.name("firstName");
	By surname = By.id("usernamereg-lastName");
	By email = By.name("yid");
	By password = By.id("usernamereg-password");
	By mobile = By.name("phone");
	By month = By.id("usernamereg-month");
	By day = By.id("usernamereg-day");
	By year = By.name("yyyy");
	By gender = By.name("freeformGender");
	
	
	public void yahooDetails() throws Exception {
		driver.findElement(signin).click();
		Thread.sleep(3000);
		driver.findElement(createacct).click();
		Thread.sleep(3000);
		driver.findElement(firstname).sendKeys("Kartheek");
		Thread.sleep(3000);
		driver.findElement(surname).sendKeys("Gunapati");
		Thread.sleep(3000);
		driver.findElement(email).sendKeys("kartheek.gunapati");
		Thread.sleep(3000);
		driver.findElement(password).sendKeys("987627333");
		Thread.sleep(3000);
		driver.findElement(mobile).sendKeys("9887233133");
		Thread.sleep(3000);
		WebElement e1 = driver.findElement(month);
		Select s = new Select(e1);
		s.selectByVisibleText("August");
		Thread.sleep(3000);
		driver.findElement(day).sendKeys("23");
		Thread.sleep(3000);
		driver.findElement(year).sendKeys("1993");
		Thread.sleep(3000);
		driver.findElement(gender).sendKeys("Male");
	}
		

}
