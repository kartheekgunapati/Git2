package Yahoo;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.ui.Select;

public class Pagefactorylibrary {
	
	WebDriver driver;
	
	public Pagefactorylibrary(WebDriver driver) {
		this.driver=driver;
	}
	
	@FindBy(xpath="//span[text()='Sign in']") WebElement signin;
	@FindBy(how=How.ID,using="createacc") WebElement createacct;
	@FindBy(id="usernamereg-firstName") WebElement firstname;
	@FindBy(how=How.NAME,using="lastName") WebElement surname;
	@FindBy(name="yid") WebElement email;
	@FindBy(how=How.NAME,using="password") WebElement password;
	@FindBy(id="usernamereg-phone") WebElement phone;
	@FindBy(how=How.NAME,using="mm") WebElement month;
	@FindBy(name="dd") WebElement day;
	@FindBy(how=How.ID,using="usernamereg-year") WebElement year;
	@FindBy(name="freeformGender") WebElement gender;
	
	public void detailsFill() throws Exception {
		signin.click();
		Thread.sleep(3000);
		createacct.click();
		Thread.sleep(3000);
		firstname.sendKeys("Kartheek");
		Thread.sleep(3000);
		surname.sendKeys("Reddy");
		Thread.sleep(3000);
		email.sendKeys("kartheek.reddy");
		Thread.sleep(3000);
		password.sendKeys("Test@1234");
		Thread.sleep(3000);
		phone.sendKeys("9988776655");
		Thread.sleep(3000);
		Select s1 = new Select(month);
		s1.selectByVisibleText("August");
		Thread.sleep(3000);
		day.sendKeys("23");
		Thread.sleep(3000);
		year.sendKeys("1990");
		Thread.sleep(3000);
		gender.sendKeys("Male");
		
	}
	

}
