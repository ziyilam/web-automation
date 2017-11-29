package config;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.ArrayList;
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
	private static int iNameCount = 1;
	private static String sCol;
	private static boolean bTryLoop = true;
	private static boolean bChangeSheet = false;
	private static String sSheetName;
	private static String sURL;

	private static void splitString(String sInput, int iSplit) {
		try {

			// split the words according to the word count. plus 1 to count start from 1
			// instead of 0
			int iSplit2 = iSplit + 1;
			logger.info("splitString|split sentence into: [" + iSplit + "]");
			// logger.info("iSplit2: [" + iSplit2 + "]");
			aWords = sInput.split("\\s", iSplit2);
			for (String w : aWords) {
				logger.info("splitString|split string: [" + w + "]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|splitString. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	private static void getAllKindsOfText(String sObjectLocator) {
		try {
			// clear old string in the memory
			DriverScript.sCompareText = "";
			String sText = driver.findElement(By.xpath(sObjectLocator)).getText();
			String sValue = driver.findElement(By.xpath(sObjectLocator)).getAttribute("value");
			String sPlaceholder = driver.findElement(By.xpath(sObjectLocator)).getAttribute("placeholder");
			String sfieldTitle = driver.findElement(By.xpath(sObjectLocator)).getAttribute("field-title");
			int iType = 0;
			if (sText.length() != 0) {
				iType = 0;
				DriverScript.sCompareText = sText;
				logger.info("getAllKindsOfText|Text grabbed: [" + sText + "]" + " iType: [" + iType + "]");
			} else if (sValue.length() != 0) {
				iType = 1;
				DriverScript.sCompareText = sValue;
				logger.info("getAllKindsOfText|Value grabbed: [" + sValue + "]" + " iType: [" + iType + "]");
			} else if (sPlaceholder.length() != 0) {
				iType = 2;
				DriverScript.sCompareText = sPlaceholder;
				logger.info("getAllKindsOfText|Placeholder grabbed: [" + sPlaceholder + "]" + " iType: [" + iType + "]");
			} else if (sfieldTitle.length() != 0) {
				iType = 2;
				DriverScript.sCompareText = sfieldTitle;
				logger.info("getAllKindsOfText|fieldTitle grabbed: [" + sfieldTitle + "]" + " iType: [" + iType + "]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|getAllKindsOfText. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;

		}
	}

	public void tryVerify(String sObjectLocator, String sTestData, String sAdditionalRequest) {

		try {
			Boolean bNot = false;
			logger.info("[tryVerify|Verifing text]");

			// get the text to verify
			getAllKindsOfText(sObjectLocator);

			// choose which word to verify
			splitAndChoose(sAdditionalRequest);

			logger.info("[tryVerify|checking for the word 'not']");
			String sInput = sTestData;
			splitString(sInput, 2);
			// String[] aWords = sInput.split("\\s",2);
			for (String w : aWords) {
				logger.info("tryVerify|checking for 'not' word: [" + w + "]");
				if (w.equalsIgnoreCase("not")) {
					bNot = true;
					break;
				} else {
					logger.info("[tryVerify|'not' word does't exist]");
				}
			}
			if (bNot == false) {
				if (DriverScript.sCompareText.equalsIgnoreCase(sTestData)) {
					logger.info(
							"tryVerify|Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + sTestData + "]");
					// logger.info("Value is: [" + sValue + "] compared with expected: [" +
					// sTestData + "]");
					logger.info("[tryVerify|Text verified to be the SAME]");
				} else {
					DriverScript.bResult = false;
					logger.info(
							"tryVerify|Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + sTestData + "]");
					// logger.info("Value is: [" + sValue + "] compared with expected: [" +
					// sTestData + "]");
					logger.info("[tryVerify|Text verified NOT the same]");
				}
			} else {
				if (!DriverScript.sCompareText.equalsIgnoreCase(aWords[1])) {
					logger.info(
							"tryVerify|Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + aWords[1] + "]");
					// logger.info("Value is: [" + sValue + "] compared with expected: [" +
					// aWords[1] + "]");
					logger.info("[tryVerify|Text verified NOT the same]");
				} else {
					DriverScript.bResult = false;
					logger.info(
							"tryVerify|Text is: [" + DriverScript.sCompareText + "] compared with expected: [" + aWords[1] + "]");
					// logger.info("Value is: [" + sValue + "] compared with expected: [" +
					// aWords[1] + "]");
					logger.info("[tryVerify|Text verified to be the SAME]");
				}
			}

		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryVerify. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void getValueNSetCell(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[getValueNSetCell|getting text]");
			getAllKindsOfText(sObjectLocator);
			logger.info("getValueNSetCell|Test step no.: [" + DriverScript.iCountTestStep + "]");
			logger.info("getValueNSetCell|AdditionalRequest: [" + sAdditionalRequest + "]");

			// choose which word to get
			splitAndChoose(sAdditionalRequest);

			// ExcelUtils.setCellData(DriverScript.sCompareText,
			// DriverScript.iCountTestStep, Constants.Col_TestData,
			// Constants.Sheet_TestSteps);
			// control the excel column to insert the word
			int iCellHeader = DriverScript.iCellHeaderIndex;
			ExcelUtils.setCellData(DriverScript.sCompareText, DriverScript.iCountTestData, iCellHeader,
					Constants.Sheet_TestData);
		} catch (Exception e) {
			logger.error(" * ActionKeywords|getValueNSetCell. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	private static void splitAndChoose(String sAdditionalRequest) {
		try {
			// if the AdditionalRequest is empty, get the whole word from web element
			if (sAdditionalRequest.isEmpty()) {
				// get the whole words from web element
				// no change on DriverScript.sCompareText, so do nothing
				logger.info("[splitAndChoose|get the whole words from web element]");
			} else {
				// split AdditionalRequest into 2 words
				//splitString(sAdditionalRequest, 2);
				// first word is to decide number of words to split
				//int iSplit = Integer.parseInt(aWords[0]);
				
				// second word is the word choose to set into the cell
				//int iSplit2 = Integer.parseInt(aWords[1]);
				// third word is the column name in the Stock sheet saved in aWords[2]
				// sCol = aWords[2];
				// count from 1 instead of 0
				//iSplit2 = iSplit2 - 1;

				// split the words from web element
				// always split all words separated by space
				int iSplit = -1;
				splitString(DriverScript.sCompareText, iSplit);
				int iSplit2 = Integer.parseInt(sAdditionalRequest);
				// count from 1 instead of 0
				iSplit2 = iSplit2 - 1;
				DriverScript.sCompareText = aWords[iSplit2];
				logger.info("splitAndChoose|chosen word no: [" + iSplit2 + "] word: [" + DriverScript.sCompareText + "]" );
			}

		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|splitAndChoose. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void copyAndPaste(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[copyAndPaste|copying text]");
			splitString(sObjectLocator, 2);
			int iRow, iCol;
			iRow = Integer.parseInt(aWords[0]);
			iCol = Integer.parseInt(aWords[1]);
			String sData = ExcelUtils.getCellData(iRow, iCol, Constants.Sheet_TestData);

			splitString(sTestData, 2);
			iRow = Integer.parseInt(aWords[0]);
			iCol = Integer.parseInt(aWords[1]);
			ExcelUtils.setCellData(sData, iRow, iCol, Constants.Sheet_TestData);

		} catch (Exception e) {
			logger.error(" * ActionKeywords|copyAndPaste. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void openBrowser(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[openBrowser|Opening the browser]");
			switch (sTestData.toLowerCase()) {
			case "chrome":
				// System.setProperty(Constants.Chrome_Property1, Constants.Chrome_Property2);
				ChromeDriverManager.getInstance().setup();
				driver = new ChromeDriver();
				driver.manage().window().maximize();
				break;

			case "ie":
				// System.setProperty(Constants.IE_Property1, Constants.IE_Property2);
				InternetExplorerDriverManager.getInstance().setup();
				driver = new InternetExplorerDriver();
				break;

			case "firefox":
				// System.setProperty(Constants.Firefox_Property1, Constants.Firefox_Property2);
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
				logger.warn("[openBrowser|Browser name not match]");
				break;
			}
			sURL = sObjectLocator;
			driver.get(sURL);
			driver.manage().deleteAllCookies();
			// driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(Constants.Sec_implicitlyWait, TimeUnit.SECONDS);
			driver.manage().timeouts().pageLoadTimeout(Constants.Sec_pageLoadTimeout, TimeUnit.SECONDS);
			

		} catch (Exception e) {
			logger.error(" * ActionKeywords|openBrowser. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryClick(String sObjectLocator, String sTestData, String sAdditionalRequest)
			throws InterruptedException {

		try {
			logger.info("tryClick|Clicking on ObjectLocator [" + sObjectLocator + "]");
			// Thread.sleep(10);
			WebElement wObject = driver.findElement(By.xpath(sObjectLocator));
			// waiting for another thread with duties that are understood to have time
			// requirements
			Thread.sleep(10);
			wObject.click();
			Thread.sleep(10);
			// driver.findElement(By.xpath(sObjectLocator)).click();
			// source = driver.findElement(By.xpath(sObjectLocator));
			// if(source.isEnabled())
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryClick. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryInput(String sObjectLocator, String sTestData, String sAdditionalRequest)
			throws InterruptedException {
		try {

			// driver.findElement(By.cssSelector(sObjectLocator)).sendKeys(sTestData);
			tryClick(sObjectLocator, sTestData, sAdditionalRequest);
			WebElement wObject = driver.findElement(By.xpath(sObjectLocator));
			// waiting for another thread with duties that are understood to have time
			// requirements
			Thread.sleep(10);
			logger.info("tryInput|Inputing [" + sTestData + "] into ObjectLocator: [" + sObjectLocator + "]");
			wObject.sendKeys(sTestData);
			Thread.sleep(10);

			// driver.findElement(By.xpath(sObjectLocator)).sendKeys(sTestData);
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryInput. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryClose(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[tryClose|Closing the browser]");
			switch (sTestData.toLowerCase()) {
			case "all":
				driver.quit();
				break;

			default:
				driver.close();
				break;
			}

		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryClose. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
		// System.exit(0);
	}

	public void openTab(String sObjectLocator, String sTestData, String sAdditionalRequest)
			throws InterruptedException {
		try {
			logger.info("[openTab|opening new tab]");
			// open new tab
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			robot.keyRelease(KeyEvent.VK_T);
			robot.delay(10);
			// switch to the new tab
			trySwitch("", "second", "");
			/*
			 * for(String handle : driver.getWindowHandles()) {
			 * driver.switchTo().window(handle); logger.info(" windowHandle: [" + handle +
			 * "]"); }
			 */
			// ArrayList<String> sTab = new ArrayList<String>(driver.getWindowHandles());
			// driver.switchTo().window(sTab.get(1));
			// to navigate to new URl in the new tab
			driver.get(sObjectLocator);
		} catch (AWTException e) {
			logger.error(" * ActionKeywords|openTab. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}

	}

	public void trySwitch(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			for (String handle : driver.getWindowHandles()) {
				// driver.switchTo().window(handle);
				logger.info("trySwitch|windowHandle: [" + handle + "]");
			}
			ArrayList<String> sTab = new ArrayList<String>(driver.getWindowHandles());
			int iHandle = 0;
			switch (sTestData.toLowerCase()) {
			case "first":
				iHandle = 0;
				break;
			case "second":
				iHandle = 1;
				break;
			case "third":
				iHandle = 2;
				break;
			default:
				iHandle = 0;
				break;
			}
			// int iHandle = Integer.parseInt(sTestData);
			driver.switchTo().window(sTab.get(iHandle));
			// waiting for another thread with duties that are understood to have time
			// requirements
			Thread.sleep(30);
			logger.info("trySwitch|switch to Handle: [" + iHandle + "]");
		} catch (Exception e) {
			logger.error(" * ActionKeywords|trySwitch. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryPopup(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			Alert alert = driver.switchTo().alert();
			switch (sTestData.toLowerCase()) {
			// accept or dismiss popup
			case "accept":
				alert.accept();
				break;
			case "dismiss":
				alert.dismiss();
				break;
			// verify its contents
			default:
				DriverScript.sCompareText = alert.getText();
				if (DriverScript.sCompareText.equalsIgnoreCase(sTestData)) {
					logger.info("tryPopup|Text is: [" + DriverScript.sCompareText + "]" + " compared with expected: ["
							+ sTestData + "]");
					logger.info("[tryPopup|Text is verified]");
				} else {
					DriverScript.bResult = false;
					logger.info("tryPopup|Text is: [" + DriverScript.sCompareText + "]" + " compared with expected: ["
							+ sTestData + "]");
					logger.info("[tryPopup|Text is not the same]");
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
			logger.info("tryNavigate|Navigating: " + sTestData + "]");
			switch (sTestData.toLowerCase()) {

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
			logger.info("[dragAndDrop|drag and drop]");
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
			logger.info("selectDropDown|Selecting dropdown item: " + sTestData + "]");
			int iDropDownIndex = Integer.parseInt(sTestData);
			(new Select(driver.findElement(By.xpath(sObjectLocator)))).selectByIndex(iDropDownIndex);
		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|selectDropDown. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void verifyURL(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("verifyURL|verifying URL: " + sObjectLocator + "]");
			DriverScript.sCompareText = driver.getCurrentUrl();
			if (DriverScript.sCompareText.equals(sObjectLocator)) {
				logger.info(
						"verifyURL|URL is: [" + DriverScript.sCompareText + "] compared with expected: [" + sObjectLocator + "]");
				logger.info("[verifyURL|URL verified]");

			} else {
				DriverScript.bResult = false;
				logger.info(
						"verifyURL|URL is: [" + DriverScript.sCompareText + "] compared with expected: [" + sObjectLocator + "]");
				logger.info("[verifyURL|Text not the same]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|verifyURL. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}

	}

	public void trySlide(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[trySlide|sliding]");
			source = driver.findElement(By.xpath(sObjectLocator));

			int iXcoordinate = source.getLocation().getX();
			int iYcoordinate = source.getLocation().getY();
			logger.info("trySlide|x1, y1: [" + iXcoordinate + "]" + ", [" + iYcoordinate + "]");

			int iOffset = Integer.parseInt(sTestData);
			int iXoffset = iXcoordinate + iOffset;
			logger.info("trySlide|x before offset, x after offset: [" + iXcoordinate + "]" + ", [" + iXoffset + "]");
			(new Actions(driver)).dragAndDropBy(source, iXoffset, iYcoordinate).perform();
		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|trySlide. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryScroll(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			logger.info("[tryScroll|scrolling]");
			JavascriptExecutor js = (JavascriptExecutor) driver;
			source = driver.findElement(By.xpath(sObjectLocator));
			int iXcoordinate = source.getLocation().getX();
			int iYcoordinate = source.getLocation().getY();
			js.executeScript("window.scrollBy(" + iXcoordinate + "," + iYcoordinate + ")", "");
		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryScroll. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void waitUntil(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 60);
			switch (sTestData) {
			case "clickable":
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath(sObjectLocator)));
				logger.info("[waitUntil|sleep until clickable]");
				break;
			case "sleep":
				if (sAdditionalRequest.isEmpty()) {
					logger.info("waitUntil|sleep for: [2000] msec");
					Thread.sleep(2000);
				} else {
					int iSleep = Integer.parseInt(sAdditionalRequest);
					// turn second into mili seconds
					iSleep = iSleep * 1000;
					logger.info("waitUntil|sleep for: [" + iSleep + "] msec");
					Thread.sleep(iSleep);
				}

				break;
			default:
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(sObjectLocator)));
				logger.info("[waitUntil|sleep until presence of element]");
				break;
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|waitUntil. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	public void tryLoop(String sObjectLocator, String sTestData, String sAdditionalRequest) {
		int iSheet = 0;
		try {
			// read excel to get all stock names
			do {
				// change to another sheet
				if (bChangeSheet) {
					iSheet++;
					iNameCount = 1;
				}

				switch (iSheet) {
				case 0:
					sSheetName = Constants.Sheet_Stk0;
					bChangeSheet = false;
					break;
				case 1:
					sSheetName = Constants.Sheet_Stk1;
					bChangeSheet = false;
					break;
				case 2:
					sSheetName = Constants.Sheet_Stk2;
					bChangeSheet = false;
					break;
				case 3:
					sSheetName = Constants.Sheet_Stk3;
					bChangeSheet = false;
					break;
				case 4:
					sSheetName = Constants.Sheet_Stk4;
					bChangeSheet = false;
					break;
				case 5:
					sSheetName = Constants.Sheet_Stk5;
					bChangeSheet = false;
					break;
				case 6:
					sSheetName = Constants.Sheet_Stk6;
					bChangeSheet = false;
					break;
				case 7:
					sSheetName = Constants.Sheet_Stk7;
					bChangeSheet = false;
					break;
				case 8:
					sSheetName = Constants.Sheet_Stk8;
					bChangeSheet = false;
					break;
				case 9:
					sSheetName = Constants.Sheet_Stk9;
					bChangeSheet = false;
					break;

				default:
					bTryLoop = false;
					logger.info("[tryLoop|all stock sheet completed]");
					logger.info("tryLoop|continue loop (true/false): [" + bTryLoop + "]");
					bChangeSheet = false;
					break;
				}
				logger.info("tryLoop|sheet number: [" + iSheet + "]");
				logger.info("tryLoop|sheet name: [" + sSheetName + "]");

				// clear old remarks
				if (!ExcelUtils.getCellData(iNameCount, Constants.Col_Remark, sSheetName).isEmpty()) {
					logger.info("tryLoop|[clearing old remarks]");
					ExcelUtils.setCellData("", iNameCount, Constants.Col_Remark, sSheetName);
				}

				// skip execute if EPS column is already filled
				String sEPS = ExcelUtils.getCellData(iNameCount, Constants.Col_EPS, sSheetName);
				logger.info("tryLoop|EPS column value: [" + sEPS + "]");
				if (sEPS.isEmpty() && bTryLoop) {
					getAndInput(sObjectLocator, sTestData, sAdditionalRequest);

					getAndPaste(sObjectLocator, sTestData, sAdditionalRequest);
					logger.info("tryLoop|continue loop (true/false): [" + bTryLoop + "]");
					// go back to the search page
					// tryNavigate(sObjectLocator, "back", sAdditionalRequest);
					driver.get(sURL);
					logger.info("tryLoop|back to page: [" + sURL + "]");
				} else {
					iNameCount++;
					// get stock name from excel, if empty then quit
					String sData = ExcelUtils.getCellData(iNameCount, Constants.Col_stk_name, sSheetName);
					logger.info("tryLoop|increased name count to: [" + iNameCount + "] and got stock name of: [" + sData
							+ "]");
					if (sData.isEmpty()) {
						// bTryLoop = false;
						bChangeSheet = true;
					}
				}
			} while (bTryLoop);
			logger.info("[tryLoop|break tryLoop]");

		} catch (Exception e) {
			logger.error(" * ActionKeywords|tryLoop. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

	private void getAndInput(String sObjectLocator, String sTestData, String sAdditionalRequest) {

		try {

			// get name from excel
			logger.info("getAndInput|getting stock name no.: [" + iNameCount + "]");
			String sData = ExcelUtils.getCellData(iNameCount, Constants.Col_stk_name, sSheetName);
			sData = sData.toUpperCase();
			logger.info("getAndInput|stock name: [" + sData + "]");

			// input the name into web element
			// change to another sheet if this sheet has no more stock name
			if (sData.isEmpty()) {
				// bTryLoop = false;
				bChangeSheet = true;
				logger.info("[getAndInput|change to another sheet]");

			} else {
				tryInput(sObjectLocator, sData, sAdditionalRequest);
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				robot.delay(10);
			}

		} catch (Exception e) {
			logger.error(" * ActionKeywords|getAndInput. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}

	}

	private void getAndPaste(String sObjectLocator, String sTestData, String sAdditionalRequest) throws Exception {
		int iCol = 0;
		try {
			// execute only when the loop is still in the same sheet
			if (bTryLoop && (!bChangeSheet)) {
				// get data from web element
				getAllKindsOfText(sTestData);
				// choose which word to paste
				splitAndChoose2(sAdditionalRequest);

				// paste the data to the excel
				logger.info("getAndPaste|column chosen: [" + sCol + "]");
				switch (sCol) {
				case "EPS":
					iCol = Constants.Col_EPS;
					break;

				default:
					iCol = Constants.Col_EPS;
					break;
				}
				ExcelUtils.setCellData(DriverScript.sCompareText, iNameCount, iCol, sSheetName);
				// increase the count after it has been pasted into the excel
				iNameCount++;
				logger.info("getAndPaste|name count increased to: [" + iNameCount + "]");
			} else {
				logger.info("[getAndPaste|skip getAndPaste]");
			}
		} catch (Exception e) {
			logger.error(" * ActionKeywords|getAndPaste. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
			// increase the count and continue the rest of the list if the stock name does
			// not exist
			// and wtite a remark in the excel
			ExcelUtils.setCellData("*****", iNameCount, Constants.Col_Remark, sSheetName);
			iNameCount++;

		}
	}

	private static void splitAndChoose2(String sAdditionalRequest) {
		try {
			// split AdditionalRequest into 3 words
			splitString(sAdditionalRequest, 3);
			// first word is to decide number of words to split
			int iSplit = Integer.parseInt(aWords[0]);
			// second word is the word choose to set into the cell
			int iSplit2 = Integer.parseInt(aWords[1]);
			// third word is the column name in the Stock sheet saved in aWords[2]
			sCol = aWords[2];
			// count from 1 instead of 0
			iSplit2 = iSplit2 - 1;

			// split the words from web element
			splitString(DriverScript.sCompareText, iSplit);

			DriverScript.sCompareText = aWords[iSplit2];
			logger.info("splitAndChoose2|chosen word no: [" + iSplit2 + "]");

		} catch (NumberFormatException e) {
			logger.error(" * ActionKeywords|splitAndChoose2. Exception Message - " + e.getMessage());
			DriverScript.bResult = false;
		}
	}

}
