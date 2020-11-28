package Yahoo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

public class Configfile {
	
	Properties pro;
	
	public Configfile() throws Exception {
		try {
			File src = new File("C:\\Users\\msuser1\\eclipse-workspace\\Git Project\\Configuration\\config.property");
			FileInputStream fis = new FileInputStream(src);
			pro = new Properties();
			pro.load(fis);
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		}
	}
	
	public String getChrome() {
		String chrome = pro.getProperty("ChormeDriver");
		return chrome;
	}
	
	public String getFirefox() {
		String firefox = pro.getProperty("FirefoxDriver");
		return firefox;
	}

	public String getfacebook() {
		String facebook = pro.getProperty("FBUrl");
		return facebook;
	}
	
	public String getyahoo() {
		String yahoo = pro.getProperty("YahooUrl");
		return yahoo;
	}
	
	public String getAbhibus() {
		String  abhibus= pro.getProperty("AbhibusUrl");
		return abhibus;
	}
	
	public String getRedbus() {
		String redbus = pro.getProperty("RedbusUrl");
		return redbus;
	}
	
	public String getGoogle() {
		String google = pro.getProperty("GoogleUrl");
		return google;
	}
	
	public String getAmazon() {
		String  amazon = pro.getProperty("AmazonUrl");
		return amazon;
	}
	
	public String getFlipkart() {
		String flipkart = pro.getProperty("FlipkartUrl");
		return flipkart;
	}
	
}
