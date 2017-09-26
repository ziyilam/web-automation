package config;

import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import executionEngine.DriverScript;
import io.github.bonigarcia.wdm.ChromeDriverManager;
import io.github.bonigarcia.wdm.FirefoxDriverManager;
import io.github.bonigarcia.wdm.InternetExplorerDriverManager;
import utility.ExcelUtils;

public class ActionKeywords {
	public static WebDriver driver;
	private static WebElement source;
	private static WebElement target;
	private static final Logger logger = LogManager.getLogger(ActionKeywords.class.getName());
	private static String[] aWords;
	
	private static void splitString (String sInput, int iSplit) {
		try {
			//int iSplit = Integer.parseInt(sSplit);
			aWords = sInput.split("\\s", iSplit);
			for(String w:aWords) {
				logger.info("w: " + w);
				}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|splitString. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	private static void getAllKindsOfText (String sObjectLocator) {
		try {
			String sText = driver.findElement(By.xpath(sObjectLocator)).getText();
			String sValue = driver.findElement(By.xpath(sObjectLocator)).getAttribute("value");
			String sPlaceholder = driver.findElement(By.xpath(sObjectLocator)).getAttribute("placeholder");
			String sfieldTitle = driver.findElement(By.xpath(sObjectLocator)).getAttribute("field-title");
			int iType = 0;
			if (sText.length()!=0) {
				iType = 0;
				DriverScript.sCompareText = sText;
				logger.info("Text: [" + sText + "]" + " iType: [" + iType + "]");
			} else if (sValue.length()!=0) {
				iType = 1;
				DriverScript.sCompareText = sValue;
				logger.info("Value: [" + sValue + "]" + " iType: [" + iType + "]");
			} else if (sPlaceholder.length()!=0){
				iType = 2;
				DriverScript.sCompareText = sPlaceholder;
				logger.info("Placeholder: [" + sPlaceholder + "]" + " iType: [" + iType + "]");
			} else if (sfieldTitle.length()!=0){
				iType = 2;
				DriverScript.sCompareText = sfieldTitle;
				logger.info("fieldTitle: [" + sfieldTitle + "]" + " iType: [" + iType + "]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|getAllKindsOfText. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void tryVerify(String sObjectLocator, String sTestData, String sAdditionalRequest) {

		try {
			Boolean bNot=false;
			logger.info("Action......Try Verify text");
			getAllKindsOfText (sObjectLocator);
					
			String sInput = sTestData;
			splitString(sInput, 2);
			//String[] aWords = sInput.split("\\s",2);
			for(String w:aWords) {
				logger.info("w: [" + w + "]");
				if(w.equalsIgnoreCase("not")) {
					bNot=true;
					break;
				}
			}
			if(bNot==false) {
				if (DriverScript.sCompareText.equalsIgnoreCase(sTestData)) {
					logger.info("Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + sTestData + "]");
					//logger.info("Value is: [" + sValue + "] compared with expected: [" + sTestData + "]");
					logger.info(" Text verified to be the SAME ");
				} else {
					DriverScript.bResult = false;
					logger.info("Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + sTestData + "]");
					//logger.info("Value is: [" + sValue + "] compared with expected: [" + sTestData + "]");
					logger.info("Text verified NOT the same");
				}
			} else {
				if (!DriverScript.sCompareText.equalsIgnoreCase(aWords[1])) {
					logger.info("Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + aWords[1] + "]");
					//logger.info("Value is: [" + sValue + "] compared with expected: [" + aWords[1] + "]");
					logger.info(" Text verified NOT the same ");
				} else {
					DriverScript.bResult = false;
					logger.info("Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + aWords[1] + "]");
					//logger.info("Value is: [" + sValue + "] compared with expected: [" + aWords[1] + "]");
					logger.info(" Text verified to be the SAME ");
				}
			}
			

		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryVerify. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void getValueNSetCell(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			getAllKindsOfText (sObjectLocator);
			/*String sText = driver.findElement(By.xpath(sObjectLocator)).getAttribute("value");
			logger.info("text: " + sText);*/
			logger.info("iTestcase: [" + DriverScript.iCountTestStep + "]");
			//splitString (DriverScript.sCompareText, 3);
			//DriverScript.sCompareText = aWords[2];
			logger.info("sAdditionalRequest: [" + sAdditionalRequest + "]");
			int iSplit = Integer.parseInt(sAdditionalRequest);
			logger.info("iSplit: [" + iSplit + "]");
			splitString (DriverScript.sCompareText,iSplit);
			int iSplit2 = iSplit-1;
			logger.info("iSplit2: [" + iSplit2 + "]");
			DriverScript.sCompareText = aWords[iSplit2];
			//ExcelUtils.setCellData(DriverScript.sCompareText, DriverScript.iCountTestStep, Constants.Col_TestData, Constants.Sheet_TestSteps);
			int iCellHeader = DriverScript.iCellHeaderIndex;
			//int iCellHeader = DriverScript.iCellHeaderIndex + 1;
			ExcelUtils.setCellData(DriverScript.sCompareText, DriverScript.iCountTestData, iCellHeader, Constants.Sheet_TestData);
		} catch (Exception e) {
			logger.error(" * ActionKeywords|getValueNSetCell. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void openBrowser(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			switch (sTestData.toLowerCase()) {
			case "chrome":
				//System.setProperty(Constants.Chrome_Property1, Constants.Chrome_Property2);
				ChromeDriverManager.getInstance().setup();
				driver = new ChromeDriver();
				driver.manage().window().maximize();
				break;

			case "ie":
				//System.setProperty(Constants.IE_Property1, Constants.IE_Property2);
				InternetExplorerDriverManager.getInstance().setup();
				driver = new InternetExplorerDriver();
				break;

			case "firefox":
				//System.setProperty(Constants.Firefox_Property1, Constants.Firefox_Property2);
				FirefoxDriverManager.getInstance().setup();
				driver = new FirefoxDriver();
				break;

			case "safari":
				driver = new SafariDriver();
				break;
				
			case "htmlunit":
				driver = new HtmlUnitDriver();
				break;

			default:
				logger.warn("Browser name not match");
				break;
			}

			driver.get(sObjectLocator);
			driver.manage().deleteAllCookies();
			// driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
			logger.info("Action......Opening the browser");

		} catch (Exception e) {
			logger.error(" * ActionKeywords|openBrowser. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryClick(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		
			try {
				logger.info("Action......Clicking on ObjectLocator [" + sObjectLocator + "]");
				driver.findElement(By.xpath(sObjectLocator)).click();
				//source = driver.findElement(By.xpath(sObjectLocator));
				//if(source.isEnabled())
			} catch (Exception e) {
				logger.error(" * ActionKeywords|tryClick. Exception Message - " + e.getMessage());
				DriverScript.bResult = false;
			}
	}

	public void tryInput(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("Action......Input the text into ObjectLocator: [" + sObjectLocator + "]");
			// driver.findElement(By.cssSelector(sObjectLocator)).sendKeys(sTestData);
			driver.findElement(By.xpath(sObjectLocator)).sendKeys(sTestData);
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryInput. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryClose(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("Action......Closing the browser");
			driver.close();
			// driver.quit();
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryClose. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
		// System.exit(0);
	}
	
	public void trySwitch(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			for(String handle : driver.getWindowHandles()) {
				driver.switchTo().window(handle);
				logger.info(" windowHandle: [" + handle + "]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|trySwitch. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void tryPopup(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			Alert alert = driver.switchTo().alert();
			switch(sTestData.toLowerCase()) {
			//accept or dismiss popup
			case "accept":
				alert.accept();
				break;
			case "dismiss":
				alert.dismiss();
				break;
			//verify its contents
			default:
				DriverScript.sCompareText = alert.getText();
				if(DriverScript.sCompareText.equalsIgnoreCase(sTestData)) {
					logger.info("Text is: [" + DriverScript.sCompareText + "]" + " compared with expected: [" + sTestData + "]");
					logger.info("Text is verified");
				}else {
					DriverScript.bResult = false;
					logger.info("Text is: [" + DriverScript.sCompareText + "]" + " compared with expected: [" + sTestData + "]");
					logger.info("Text is not the same");
				}
				break;
			
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryPopup. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void tryNavigate(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			switch(sTestData.toLowerCase()) {
			
			// move back a single "item" in the browser's history
			case "back":
				driver.navigate().back();
				break;
			// move a single "item" forward in the browser's history
			case "forward":
				driver.navigate().forward();
				break;
			// refresh the current page
			case "refresh":
				driver.navigate().refresh();
				break;
			// load a new web page in the current browser window
			default:
				driver.navigate().to(sTestData);
				break;
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryNavigate. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void dragAndDrop(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			source = driver.findElement(By.xpath(sObjectLocator));
			target = driver.findElement(By.xpath(sTestData));
			(new Actions(driver)).dragAndDrop(source, target).perform();
		} catch (Exception e) {
			logger.error(" * ActionKeywords|DragAndDrop. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
		
	}
	
	public void selectDropDown(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			int iDropDownIndex = Integer.parseInt(sTestData);
			(new Select(driver.findElement(By.xpath(sObjectLocator)))).selectByIndex(iDropDownIndex);
			logger.info(" selecting dropdown item ");
		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|selectDropDown. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void verifyURL(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			DriverScript.sCompareText = driver.getCurrentUrl();
			if(DriverScript.sCompareText.equals(sObjectLocator)) {
				logger.info("URL is: [" + DriverScript.sCompareText + "] compared with expected: [" + sObjectLocator + "]");
				logger.info("URL verified");
				
			}else {
				DriverScript.bResult = false;
				logger.info("URL is: [" + DriverScript.sCompareText + "] compared with expected: [" + sObjectLocator + "]");
				logger.info("Text not the same");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|verifyURL. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
		
	}
	
	public void trySlide(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			source = driver.findElement(By.xpath(sObjectLocator));
			
			int iXcoordinate = source.getLocation().getX();
			int iYcoordinate = source.getLocation().getY();
			logger.info("x1, y1: [" + iXcoordinate + "]" + ", [" + iYcoordinate + "]");
			
			int iOffset = Integer.parseInt(sTestData);
			int iXoffset = iXcoordinate + iOffset;
			logger.info("x before offset, x after offset: [" + iXcoordinate + "]" + ", [" + iXoffset + "]");
			(new Actions(driver)).dragAndDropBy(source, iXoffset, iYcoordinate).perform();
		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|trySlide. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void tryScroll(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			JavascriptExecutor js = (JavascriptExecutor)driver;
			source = driver.findElement(By.xpath(sObjectLocator));
			int iXcoordinate = source.getLocation().getX();
			int iYcoordinate = source.getLocation().getY();
			js.executeScript("window.scrollBy("+iXcoordinate+","+iYcoordinate+")","");
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryScroll. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}
	
	public void waitUntil(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 60);
			switch(sTestData) {
			case "clickable":
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath(sObjectLocator)));
				break;
			default:
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(sObjectLocator)));
				break;
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|waitUntil. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

}
