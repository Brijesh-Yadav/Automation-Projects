package democlasses;

import java.awt.AWTException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import configuration.Resourse_path;

public class BAFileDownload {
	
	public static void main(String [] args) throws AWTException, InterruptedException{
		
		
		String chromedriverserver = Resourse_path.chrome_driver_path;
		System.setProperty("webdriver.chrome.driver",chromedriverserver);
		WebDriver driver = new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.get("http://ausyd16as13v/BA2Web_WIP/ba2admin/Home/tabid/121/Default.aspx");
		
		driver.findElement(By.id("dnn_ctr662_LoginDefault_SignIn1_txtUsername")).sendKeys("pratibhaseth");
		driver.findElement(By.id("dnn_ctr662_LoginDefault_SignIn1_txtPassword")).sendKeys("Password123$");
		
		String Lable_color =  driver.findElement(By.className("fieldTitle")).getCssValue("color");
		String Lable_background_color =  driver.findElement(By.className("fieldTitle")).getCssValue("background-color");	
		String Lable_font_weight = driver.findElement(By.className("fieldTitle")).getCssValue("font-weight");
		String Lable_font_fmaily = 	driver.findElement(By.className("fieldTitle")).getCssValue("font-family");
		
		System.out.println("Properties - Lable_color :: "+Lable_color+" , Lable_background_color :: "+Lable_background_color+" , Lable_font_weight :: "+Lable_font_weight +" , Lable_font_fmaily :: "+Lable_font_fmaily );
		
		/*
		driver.findElement(By.id("dnn_ctr662_LoginDefault_SignIn1_cmdLogin")).click();
	
		String url2 = "http://ausyd16as13v/BA2Web_WIP/ba2admin/Payment/GenerateFilesNew/tabid/6754/ctl/PaymentGenerateNew/mid/1082/Default.aspx";
		
		driver.navigate().to(url2);
		
		driver.findElement(By.linkText("Payment_Summary.xlsx")).click();
		
		Thread.sleep(5*1000);
		
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		*/
		driver.close();
		
	}

}
