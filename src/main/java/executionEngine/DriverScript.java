package executionEngine;

import java.lang.reflect.Method;
import java.util.ArrayList;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import config.ActionKeywords;
import config.Constants;
import utility.ExcelUtils;

public class DriverScript {
	public static ActionKeywords actionKeywords;
	public static String sActionKeyword;
	public static String sObjectLocator;
	public static String sTestData;
	public static String sAdditionalRequest;
	public static String sTestDataItem;
	public static String sTestStepID;
	public static String sTestCaseID;
	public static String sRunMode;
	public static String sCompareText;
	public static boolean bResult;
	public static Method method[];
	public static ArrayList<String> alCellHeader;
	public static int iTotalTestCases;
	public static int iCellHeaderIndex;
	public static int iStartTestData;
	public static int iLastTestData;
	public static int iCountTestData;
	public static int iStartTestStep;
	public static int iLastTestStep;
	public static int iCountTestStep;
	public static int iCountCol;
	public static int iTestcase;

	private static final Logger logger = LogManager.getLogger(DriverScript.class.getName());

	public DriverScript() throws NoSuchMethodException, SecurityException {
		actionKeywords = new ActionKeywords();
		method = actionKeywords.getClass().getMethods();
		alCellHeader = new ArrayList<>();
	}

	public static void main(String[] args) throws Exception {

		ExcelUtils.setExcelFile(Constants.Path_TestData);
		DriverScript startEngine = new DriverScript();
		startEngine.execute_TestCase();

	}

	private void execute_TestCase() throws Exception {
		logger.info("\n\n---------------------------------------   execute_TestCase|START  ---------------------------------------\n\n\n");
		iTotalTestCases = ExcelUtils.getRowCount(Constants.Sheet_TestCases);
		logger.info("execute_TestCase|Total TestCases: [" + iTotalTestCases + "]");
		// record the column available in the testData sheet
		record_HeaderName();

		// Loop from test case no.1 to the last test case
		for (iTestcase = 1; iTestcase <= iTotalTestCases; iTestcase++) {
			// Every new Test case the bResult is reset to true
			bResult = true;
			logger.info("execute_TestCase|TestCase no.: [" + iTestcase + "]");
			sTestCaseID = ExcelUtils.getCellData(iTestcase, Constants.Col_TestCaseID, Constants.Sheet_TestCases);
			logger.info("execute_TestCase|TestCaseID: [" + sTestCaseID + "]");
			sRunMode = ExcelUtils.getCellData(iTestcase, Constants.Col_RunMode, Constants.Sheet_TestCases);
			logger.info("\n\nexecute_TestCase|sRunMode: [" + sRunMode + "]"
					+ "\n\n\n");
			// clear previous data in Test Case Results column
			if(!ExcelUtils.getCellData(iTestcase, Constants.Col_CaseResults, Constants.Sheet_TestCases).isEmpty()) {
			logger.info("execute_TestCase|Clearing old result data for TestCase: [" + sTestCaseID + "]");
			ExcelUtils.setCellData(Constants.KEYWORD_EMPTY, iTestcase, Constants.Col_CaseResults,
					Constants.Sheet_TestCases);
			}
			// clear previous data in Test Step Result column
			iStartTestStep = ExcelUtils.getRowStartWith(sTestCaseID, Constants.Col_TestCaseID,
					Constants.Sheet_TestSteps);
			iLastTestStep = ExcelUtils.getStepsCount(Constants.Sheet_TestSteps, sTestCaseID, iStartTestStep);
			if(!ExcelUtils.getCellData(iStartTestStep, Constants.Col_StepResults, Constants.Sheet_TestSteps).isEmpty()) {
			for(;iStartTestStep<=iLastTestStep;iStartTestStep++) {
				logger.info("execute_TestCase|Clearing old result data for TestStep: [" + iStartTestStep + "]");
				ExcelUtils.setCellData(Constants.KEYWORD_EMPTY, iStartTestStep, Constants.Col_StepResults,Constants.Sheet_TestSteps);
			}
			}
			// Only execute the test case with run mode equals to yes
			if (sRunMode.equalsIgnoreCase("Yes")) {
				//startTestCase
				logger.warn("\n\n-------------------             execute_TestCase|[" + sTestCaseID + "]"
						+ " TestStep begins                -------------------\n\n\n");
				iStartTestStep = ExcelUtils.getRowStartWith(sTestCaseID, Constants.Col_TestCaseID,
						Constants.Sheet_TestSteps);
				logger.info("execute_TestCase|1st TestStep at row: [" + iStartTestStep + "]");
				iLastTestStep = ExcelUtils.getStepsCount(Constants.Sheet_TestSteps, sTestCaseID, iStartTestStep);
				logger.info("execute_TestCase|Last TestStep at row: [" + iLastTestStep + "]");
				
				// if there is any testData for this testCase only calculate TestData total row
				iStartTestData = ExcelUtils.getRowStartWith(sTestCaseID, Constants.Col_TestCaseID,
						Constants.Sheet_TestData);
				iLastTestData = ExcelUtils.getStepsCount(Constants.Sheet_TestData, sTestCaseID, iStartTestData);
				logger.info("execute_TestCase|1st TestData at row: [" + iStartTestData + "]");
				logger.info("execute_TestCase|Last TestData at row: [" + iLastTestData + "]");

				// Execute for sets of different test data
				// if there is testData, loop testData, else run without test data
				// at least run one time
				check_IfGotTestData(iStartTestData, iLastTestData);
				
				for (iCountTestData = iStartTestData; iCountTestData <= iLastTestData; iCountTestData++) {
					
					// every new set of test data the bResult is reset to true
					bResult = true;
					logger.info("\n\nexecute_TestCase|Test start for new TestData"
							+ "\n\n\n");
					
					// Loop for all test steps
					for (iCountTestStep = iStartTestStep; iCountTestStep <= iLastTestStep; iCountTestStep++) {
						
						logger.info("execute_TestCase|TestStep row no.: [" + iCountTestStep + "]");
						
						
						// get all data from test step columns, need iCountTestStep
						fetch_TestSteps(iCountTestStep);
								
						// get testData from TestData Sheet
						fetch_TestData(sTestData);
						
						// perform keyword actions
						execute_Action();

						// Record test case fail and close browser, need iTestcase
						// check_TestCaseResult(bResult, iTestcase);
						if (bResult == false) {
							ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iTestcase, Constants.Col_CaseResults,
									Constants.Sheet_TestCases);
							logger.warn("\n\n......execute_TestCase|Test Case Failed for [" + sTestCaseID + "]" + "......"
									+ "\n\n\n");
							//actionKeywords.tryClose("", "", "");
							//logger.info("close browser from TestCase loop");
							break;
						}

					}
					
					// Record test case pass, need iTestcase
					if (bResult == true) {
						ExcelUtils.setCellData(Constants.KEYWORD_PASS, iTestcase, Constants.Col_CaseResults,
								Constants.Sheet_TestCases);
						logger.info("\n\n......execute_TestCase|All TestStep Completed for [" + sTestCaseID + "]" + "......\n\n\n");
					}
					// If test case fail, skip for the rest of the data set
					if (bResult == false) {
						break;
					}
					
				} 
				//endTestCase
				logger.warn("\n\n--------------------------------execute_TestCase|[" + sTestCaseID + "]" + " TestCase " + " Ended"
						+ "------------------------------------\n\n\n");
			}
		}
		logger.info("\n\n......execute_TestCase|No more TestCase with 'Yes' RunMode......\n\n"
				+ "---------------------------------------   E-N-D  ---------------------------------------" 
				+ "\n\n\n");
		
	}
	
	/*private static void check_TestCaseResult(Boolean bRsult, int iTcase) {
		
	}*/
	private static void check_IfGotTestData(int iStart, int iLast) {
		logger.info("[check_IfGotTestData|checking if there is any Test Data]");
		int iCount = iLastTestData - iStartTestData;
		logger.info("check_IfGotTestData|icount: [" + iCount + "]");
		if(iCount < 0) {
			iLastTestData = iStartTestData;
			logger.info("check_IfGotTestData|1st TestData at row: [" + iStartTestData + "]");
			logger.info("check_IfGotTestData|Last TestData at row: [" + iLastTestData + "]");
		}
	}
	private static void execute_Action() throws Exception {
		Boolean bKeyword = false;
		try {
			logger.info("[execute_Action|executing action keywords]");
			for (int i = 0; i < method.length; i++) {
				if (method[i].getName().equalsIgnoreCase(sActionKeyword)) {
					bKeyword = true;
					method[i].invoke(actionKeywords, sObjectLocator, sTestDataItem, sAdditionalRequest);
					//method[i].invoke(actionKeywords, sObjectLocator, sActionKeyword, sTestDataItem);
					logger.info("execute_Action|Executed TestData Item: [" + sTestDataItem + "]");
					if (bResult == true) {

						logger.info("execute_Action|bResult: [" + bResult + "]" + " for [" + sTestStepID + "]");
						ExcelUtils.setCellData(Constants.KEYWORD_PASS, iCountTestStep, Constants.Col_StepResults,
								Constants.Sheet_TestSteps);
						break;
					} else {
						logger.info("execute_Action|bResult: [" + bResult + "]" + " for [" + sTestStepID + "]");
						ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iCountTestStep, Constants.Col_StepResults,
								Constants.Sheet_TestSteps);
						break;
					}
				}

			}
			if (bKeyword == false) {
				logger.warn("execute_Action|No such action keyword of: [" + sActionKeyword + "]");
				bResult = false;
				logger.info("execute_Action|bResult: [" + bResult + "]" + " [TS_" + iCountTestStep + "]");
				ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iCountTestStep, Constants.Col_StepResults,
						Constants.Sheet_TestSteps);
			}

		} catch (Exception e) {
			logger.error(" * DriverScript|execute_Action. Exception message: " + e.getMessage());
		}

	}
	
	private void fetch_TestData(String sTestData) throws Exception {
		try {
			logger.info("[fetch_TestData|fetching test data value]");
			sTestDataItem = sTestData;
			logger.info(" fetch_TestData|sTestDataItem: [" + sTestDataItem + "]");
			for(String sHeader:alCellHeader) {
				
				if(sHeader.equalsIgnoreCase(sTestData)) {
					iCellHeaderIndex = alCellHeader.indexOf(sHeader);
					sTestDataItem = ExcelUtils.getCellData(iCountTestData, iCellHeaderIndex, Constants.Sheet_TestData);
					logger.info("fetch_TestData|TestData sheet column: [" + sHeader + "]");
					logger.info("fetch_TestData|Fetching TestData Item: [" + sTestDataItem + "]");
					break;
				}
			}
			
		} catch (Exception e) {
			logger.error(" * DriverScript|fetch_TestData. Exception message: " + e.getMessage());
		}
	}
	
	private void fetch_TestSteps(int iTestStepCount) throws Exception{
		try {
			logger.info("[fetch_TestSteps|fetching test steps instructions]");
			sObjectLocator = ExcelUtils.getCellData(iCountTestStep, Constants.Col_ObjectLocator,
					Constants.Sheet_TestSteps);
			sActionKeyword = ExcelUtils.getCellData(iCountTestStep, Constants.Col_ActionKeyword,
					Constants.Sheet_TestSteps);
			sTestStepID = ExcelUtils.getCellData(iCountTestStep, Constants.Col_TestStepID,
					Constants.Sheet_TestSteps);
			sTestData = ExcelUtils.getCellData(iCountTestStep, Constants.Col_TestData,
					Constants.Sheet_TestSteps);
			sAdditionalRequest = ExcelUtils.getCellData(iCountTestStep, Constants.Col_AdditionalRequest,
					Constants.Sheet_TestSteps);
	
			logger.info("fetch_TestSteps|ActionKeyword: [" + sActionKeyword + "]");
			logger.info("fetch_TestSteps|TestData: [" + sTestData + "]");
			logger.info("fetch_TestSteps|AdditionalRequest: [" + sAdditionalRequest + "]");
			
		} catch (Exception e) {
			logger.error(" * DriverScript|fetch_TestSteps. Exception message: " + e.getMessage());
		}
	}
	
	private static void record_HeaderName() throws Exception {

		try {
			logger.info("[record_HeaderName|recording header names]");
			iCountCol = ExcelUtils.getColCount(Constants.Sheet_TestData, 0);
			logger.info("record_HeaderName|TestData sheet ColCount: [" + iCountCol + "]");
			for (int i = 0; i < iCountCol; i++) {

				alCellHeader.add(ExcelUtils.getCellData(0, i, Constants.Sheet_TestData));

			}
			logger.info("record_HeaderName|Header Name for TestData Sheet: [" + alCellHeader + "]");
		} catch (Exception e) {
			logger.error(" * DriverScript|record_HeaderName. Exception message: " + e.getMessage());
		}

	}


}
