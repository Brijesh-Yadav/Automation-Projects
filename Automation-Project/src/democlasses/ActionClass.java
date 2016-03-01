package democlasses;

import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.Action;


public class ActionClass {
	
	public static void main(String[] args) {
		  WebDriver selenium = new FirefoxDriver();
		  System.out.println("Launching Browser");
		  //Opening the URL
		  selenium.get("http://www.uftHelp.com");
		  //Implicit wait for the browser to launch
		  selenium.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		  //Google Search Text-box 
		  WebElement element =selenium.findElement(By.id("gsc-i-id1"));
		  //Creating action class object
		        Actions oAction=new Actions(selenium);
		        //Creating action collection to perform numerous methods on element
		        Actions moreActions = oAction
		           .moveToElement(element)
		           .click()
		           .keyDown(element,Keys.SHIFT) //for caps
		           .sendKeys(element,"selenium");
		        
		        Action enterInCaps = moreActions.build();
		        enterInCaps.perform();
		         System.out.println("Text is entered in Captial letters"); 
		  //Closing the browser 
		  //selenium.quit();
		 }

}
