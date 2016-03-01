package driver;

import genericclasses.GenericFunctions;
import genericclasses.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import configuration.Resourse_path;

public class Driver {

	// static variables
	public static String StepNumber = "";
	public static String DriverSheetname = "";
	public static String Testcasename = "";
	public static String TestCaseID = "";
	public static String TestData_Sheetpath = null;
	public static String KeywordName = "";
	public static String SuiteName = "";
	public static String ORXMLName = "";
	public static Integer executing_row = null;
	public static String OR_ObjectName = "";
	public static String ScreenName = "";
	public static String TestStatus = "Passed";
	public static String S_No = "";
	public static String start_time = "";
	public static String End_Time = "";
	public static String HL_RSheetName = "";
	public static String HighLevel_Result_Sheetpath = "";
	public static String LowLeveL_Result_Sheetpath = "";
	public static String LowLevel_Result_Folder = "";
	public static String ORNAME_XML = "";
	public static String Snap_flag = "";
	public static String Empcsvfname = "";
	public static String Depcsvfname = "";
	public static String Histcsvfname = "";
	public static String Snap_URL = "dummy";
	public static String Parentwindow = null;
	public static String browsername = null;
	public static String BrowserType = "";
	public static String Temp_ResultSheetPath = "";

	// To store results e.g. Passed or Failed
	public static ArrayList<String> Arr_list;
	public static ArrayList<String> FL = null;
	public static TreeMap<String, String> GVmap = new TreeMap<String, String>();

	// main function
	public static void main(String[] args) {
		try {
			// csvOperations();
			executeTestCases();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	// test execution function called by main function
	public static void executeTestCases() throws IOException,
			InvalidFormatException {
		// creating log file inside log folder using timestamp
		Logger.createLogFile();
		DataFormatter objDefaultFormat = new DataFormatter();
	    
		// Log entry in Log file using the info function in Logger class
		Logger.info("Selenium version :: " + Resourse_path.selenium_version);
		// force quit the IEDriver process thread
		// GenericFunctions.killProcess("IEDriverServer");
		// Create a folder for the current date in the Downloads folder. the
		// folder with the current date
		// have subsequent subfolders with the current time as name
		GenericFunctions.createDateTimeFolder("Downloads");

		// Create an Array List Arr_List and store passed as initial value
		Arr_list = new ArrayList<String>();
		Arr_list.add("Passed");

		// Loops through the Driversheet and returns all files to be executed in
		// an Array List
		ArrayList<String> All = GenericFunctions.ReadDriverSuiteExcel();
		int len = All.size(); // number of test case files to be executed

		Logger.info("Total number of applications :: " + len);

		// Loop equal to the number of items in the Driver sheet
		for (int i = 0; i < len; i++) {

			int counter = 0;
			DriverSheetname = (String) All.get(i);// Get the name of the test
													// case sheet to be executed
													// from the array
			HL_RSheetName = DriverSheetname; // name of the sheet is stored in a
												
			// specified in driver sheet and date and time stamp to form the
			// complete Result excel sheet name
			Snap_flag = GenericFunctions.Snp_flag.get(i).toString();// get the
																	
//			Empcsvfname = GenericFunctions.empfname.get(i).toString();
//			Depcsvfname = GenericFunctions.depfname.get(i).toString();
//			Histcsvfname = GenericFunctions.hisfname.get(i).toString();
			
			browsername = GenericFunctions.Browser_Type;// get the name of the
														// browser used for
														// execution from the
														// driver sheet
			String SkipTestCase = "False";// flag to check if test case has to
											// be skipped, initial value set as

			Logger.info("Application Name - " + DriverSheetname);

			TestData_Sheetpath = Resourse_path.TestData_Sheetpath
					+ DriverSheetname + ".xlsx"; // get the full path of the
													// test case sheet to be
													// executed
			Logger.info("TestData_Sheet path:: " + TestData_Sheetpath);
						

			FileInputStream FIS = new FileInputStream(TestData_Sheetpath);// get
			

			// Excel operation
			XSSFWorkbook Wbook_obj = new XSSFWorkbook(FIS);// create workbook
															// object
			//XSSFWorkbook Wbook_obj2 = new XSSFWorkbook(FIS);
			FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) Wbook_obj);
			// This function checks the precence of the excel sheet "Test Case"
			// in the workbook object
			GenericFunctions.checkExcelSheetPresence(Wbook_obj);
			// gets the object of the Test Case worksheet
			XSSFSheet Wsheet_obj = Wbook_obj.getSheet("Test Case");

			// variables for iterating through the test case file
			String PrevTestID = null;
			String CurrentTestID = "";
			String nextTestID = "";
			XSSFRow rowobj = null;
			XSSFRow nextRowObj = null;
			XSSFRow PrevRowObj = null;
			XSSFCell Driver_obj_1, Driver_obj_2, Driver_obj_3, Driver_obj_4;
			@SuppressWarnings("unused")
			String Condition;
			String TestCaseChangeFlag = "False";
			String PrevStep_ScreenName = "";

			// Row count operation
			int Row_c = Wsheet_obj.getLastRowNum() + 1;// Total number of rows
														// in the excel sheet
			Logger.info("Total row " + Row_c);
			int introw = 0;
			// Loop through all the rows in the excel sheet and checks if there
			// is a blank row if there is a blank row then reduces the count of
			// the rows
			for (int f = 1; f < Row_c; f++) {
				boolean rw_rs = GenericFunctions.isRowEmptyInExcel(Wsheet_obj
						.getRow(f));
				if (rw_rs != true) {
					introw++;
				}
			}

			// final row count
			int Row_count = introw;
			Logger.info("After row validation for blank or null!!");
			Logger.info("Total final row " + Row_count);

			XSSFSheet Wsheet_obj1 = Wbook_obj.getSheet("TestCaseScheduler");
			// function to read the testcasescheduler sheet and return the names
			// of all test cases to be executed
			ArrayList<String> TC = GenericFunctions.ReadTestCaseSchedulerExcel(TestData_Sheetpath);
					
			int TotalTC = TC.size(); // number of test cases to be executed
			Logger.info("Total number of test cases to be executed :: "
					+ TotalTC);

			// Loop for all test cases to be executed
			for (int t = 0; t < TotalTC; t++) {

				// test case id value picked from the array list returned from
				// the ReadTestCaseScheduler method
				TestCaseID = (String) TC.get(t);

				// Loop through all the rows in the excel sheet
				for (int k = 1; k <= Row_count; k++) {

					try {

						// if last row not reached
						if (k <= Row_count) {
							// if test case has changed
							if (TestCaseChangeFlag == "True") {

								GenericFunctions.Close();
								TestCaseChangeFlag = "False";
							}
							// get the object for the current row
							rowobj = Wsheet_obj.getRow(k);
							Driver_obj_1 = rowobj.getCell(1);
							// get the step number
							Condition = Driver_obj_1.getStringCellValue();
							Driver_obj_2 = rowobj.getCell(0);
							// get the current test case id
							CurrentTestID = Driver_obj_2.getStringCellValue();
							// this matches the test case id to be executed with
							// the current test id and if
							// not matching then sets the flag for skipping test
							// case
							if (!TestCaseID.equalsIgnoreCase(CurrentTestID)) {
								SkipTestCase = "True";
							}
							// if test case to be executed and current test case
							// matches then set flag for
							// skipping test case as false and continue
							// execution
							else if (TestCaseID.equalsIgnoreCase(CurrentTestID)) {

								SkipTestCase = "False";
								// if next row is not last row then
								if ((k + 1) <= Row_count) {
									// get the object for the next row in the
									// excel sheet
									nextRowObj = Wsheet_obj.getRow(k + 1);
									// get the value for the test case id in the
									// next row
									Driver_obj_3 = nextRowObj.getCell(0);
									nextTestID = Driver_obj_3
											.getStringCellValue();
								} else
									// if it is last row the nexttestid is set
									// as null
									nextTestID = null;
								{

									if (!CurrentTestID.equals(nextTestID)) {
										// GenericFunctions.Close();
										TestCaseChangeFlag = "True";
									}
								}

								// if value of prev test id is null and last
								// test case has not been reached then
								// this condition will be true if we are
								// executing the first test case
								if (PrevTestID == null && k <= Row_count) {
									// System.out.println("PrevTestID - "+PrevTestID
									// +" == null && row "+k +"<= "+Row_count);
									Logger.info("PrevTestID - " + PrevTestID
											+ " == null && row " + k + "<= "
											+ Row_count);
									Date dte = new Date();
									DateFormat df = DateFormat
											.getTimeInstance();
									String S_time = df.format(dte);
									// get the start time for execution
									start_time = S_time;
									// start incrementing counter
									counter = counter + 1;
									S_No = Integer.toString(counter);
									PrevTestID = CurrentTestID;
								}

								// display previous test step id
								int teststep = k - 1;
								Logger.info("Test ID "
										+ CurrentTestID
										+ " : previous test step/row "
										+ teststep
										+ " status is "
										+ TestStatus
										+ " and value of next test step/row is :: "
										+ k);
								Resourse_path.running_path = DriverSheetname
										+ File.separator + CurrentTestID
										+ File.separator;

								// get the object for the previous row
								PrevRowObj = Wsheet_obj.getRow(k - 1);
								Driver_obj_4 = PrevRowObj.getCell(3);
								// get the value for the screen in the previous
								// step
								PrevStep_ScreenName = Driver_obj_4
										.getStringCellValue();
								// GenericFunctions.createDownloadFolderpath();
								// if we have not reached last row and test case
								// status is passed
								if ((!(k > Row_count))
										&& TestStatus
												.equalsIgnoreCase("passed")) {

									executing_row = k;
									// System.out.println("\n");
									Logger.info("\n");
									// System.out.println("-------------------- Executing Testcase ID "+CurrentTestID+" and test step : "+k+" ---------------------- ");
									Logger.info("-------------------- Executing Testcase ID "
											+ CurrentTestID
											+ " and test step : "
											+ k
											+ " ---------------------- ");

									// get the value for stepnum
									XSSFCell Cell_obj_0 = rowobj.getCell(1);
									String StepNum = Cell_obj_0
											.getStringCellValue();

									// get the value for test case if
									XSSFCell Cell_obj = rowobj.getCell(0);
									String Excel_Tcid = Cell_obj
											.getStringCellValue();

									// get the value for screen name to acces
									// the right OR folder
									XSSFCell Cell_obj2 = rowobj.getCell(3);
									String ORName = Cell_obj2
											.getStringCellValue();

									// get the value for the screen name
									XSSFCell Cell_obj4 = rowobj.getCell(3);
									String Screen_Name = Cell_obj4
											.getStringCellValue();

									// get the value for the keyword
									XSSFCell Cell_obj1 = rowobj.getCell(4);
									String Keyword = Cell_obj1
											.getStringCellValue();

									String ObjectNm = "";
									// get the value for first fieldname
									XSSFCell Cell_obj9 = rowobj.getCell(5);
									if (Cell_obj9 != null) {
										ObjectNm = Cell_obj9
												.getStringCellValue();
									}


									XSSFCell Cell_ORNAME = rowobj.getCell(2);
									ORNAME_XML = Cell_ORNAME
											.getStringCellValue();
									StepNumber = StepNum;
									// System.out.println(StepNumber);
									Testcasename = Excel_Tcid;
									// System.out.println(Testcasename);
									KeywordName = Keyword;
									// System.out.println(Keyword);
									SuiteName = DriverSheetname;
									ORXMLName = ORName;
									ScreenName = Screen_Name.trim();
									// get the value of the first fieldname
									OR_ObjectName = ObjectNm;
									// System.out.println(ObjectNme);

									String SheetNameCheck = "";
									String ResultSheetname = "";

									// HL_RSheetName stores the name of the test
									// case excel file
									if (!SheetNameCheck
											.equalsIgnoreCase(HL_RSheetName)) {
										ResultSheetname = HL_RSheetName + "_"
												+ Resourse_path.DateTimeStamp;
										SheetNameCheck = HL_RSheetName;
									}

									// create the complete path to the result
									// folder
									Temp_ResultSheetPath = Resourse_path.currPrjDirpath
											+ "/Results/"
											+ Driver.HL_RSheetName
											+ "_"
											+ Resourse_path.DateTimeStamp + "/";
									// create the path to the result excel file
									HighLevel_Result_Sheetpath = Temp_ResultSheetPath
											+ "Result_"
											+ ResultSheetname
											+ ".xlsx";
									// same as HighLevel_Result_Sheetpath i.e
									// excel result file
									LowLeveL_Result_Sheetpath = Temp_ResultSheetPath
											+ "Result_"
											+ ResultSheetname
											+ ".xlsx";
									// folder for storing results
									LowLevel_Result_Folder = Resourse_path.currPrjDirpath
											+ "/Results/"
											+ ResultSheetname
											+ "/";

									// Adding test data to arraylist B1
									ArrayList<String> Bl = GenericFunctions.FindTestData();
									// store B1 in FL
									FL = Bl;
									//
									// get the total number of open windows
									int wincnt = GenericFunctions.winCount();
									String act_wind = null;
									System.out.println(wincnt);
									// String act_wind =
									// GenericFunctions.fn_Data("activateWindow");
									// if screen name has changed then set flag
									// act_wind to 1
									if (!Screen_Name.equals(PrevStep_ScreenName)) {
										act_wind = "1";
									}

									// if count of open windows is more than 1
									if (wincnt > 1) {
										System.out
												.println("More than 1 window exist, so switching to another window!!");
										Logger.info("More than 1 window exist, so switching to another window!!");
										// if keyword being called is close
										if (Keyword.equalsIgnoreCase("close")) {
											// GenericFunctions.handleMultipleWindowsforClosingwindow();
										}
										// if keyword being called is other
										else if (Keyword
												.equalsIgnoreCase("other")) {
											// System.out.println("not doing anything!!");
											Logger.info("not doing anything!!");
										}
										// if keyword being called is wait
										else if (Keyword
												.equalsIgnoreCase("wait")) {
											// System.out.println("not doing anything!!");
											Logger.info("not doing anything!!");
										}
										// if keyword being called is openurl
										else if (Keyword
												.equalsIgnoreCase("openurl")) {
											// System.out.println("not doing anything!!");
											Logger.info("not doing anything!!");
										}
										// the above if else if blocks are not
										// doing anything if we have more than 1
										// window
										// open then this else block is called
										else {

											GenericFunctions
													.windowInstancehandler("pageObject");
										}
									}

									// if we have only one window open
									else if (wincnt == 1) {
										if (act_wind != null) {
											// if the screen name has changed
											if (act_wind.equalsIgnoreCase("1")) {
												// get instance of parent window
												if (Keyword
														.equalsIgnoreCase("close")) {
													// GenericFunctions.handleMultipleWindowsforClosingwindow();
												} else if (Keyword
														.equalsIgnoreCase("other")) {
													// System.out.println("not doing anything!!");
													Logger.info("not doing anything!!");
												} else if (Keyword
														.equalsIgnoreCase("wait")) {
													// System.out.println("not doing anything!!");
													Logger.info("not doing anything!!");
												} else if (Keyword
														.equalsIgnoreCase("openurl")) {
													Logger.info("not doing anything!!");
													// System.out.println("not doing anything!!");
												} else {
													GenericFunctions
															.windowInstancehandler("pageObject");
												}
											}
										}
									}

									// Block to be excuted incase Keyword is
									// Login
									if (Keyword.equalsIgnoreCase("Login")) {
										String conditional_cl = GenericFunctions
												.fn_Data("conditional_click");
										if (conditional_cl != null) {
											// checks the presence of element
											boolean sts = GenericFunctions
													.validateElementPresence(conditional_cl);
											if (sts != true) {
												GenericFunctions
														.Login(GenericFunctions
																.fn_Data("UserName"),
																GenericFunctions
																		.fn_Data("PASSWORD"));
											}
										}
										// this is called in normal conditions
										// for Login
										else {
											GenericFunctions.Login(GenericFunctions.fn_Data("UserName"),GenericFunctions.fn_Data("PASSWORD"));
													
										}

										// stores the return value from Login
										// function in a variable resulttst
										int resultst = GenericFunctions.exe_rst_status;

										reportResult(resultst,
												"User should be logged-In",
												"Login Successful",
												"User failed while login.");

									}
									// this block is executed if keyword is
									// Validate
									else if (Keyword.equalsIgnoreCase("Validate")) {
											
										// create two arraylists a1Name and
										// a1Val
										String object_list = OR_ObjectName;
										
										String [] object_arr_list = object_list.split(",");
										
										String action = GenericFunctions.fn_Data("action");
										
										if(object_arr_list.length>1){
											 GenericFunctions.validate_action_list(action);
										}else {
											if(action != null){
												 GenericFunctions.validate_action_list(action);
											}else {
												ArrayList<String> alName = new ArrayList<String>();
												ArrayList<String> alVal = new ArrayList<String>();

												// get the total number of columns
												int Cell_count = rowobj.getLastCellNum();
												String cell_val5 = null;
												// loop through all the fieldname
												// columns and store the values in all
												// fieldnames in the
												// a1Name arraylist
												for (int j = 5; j <= Cell_count; j = j + 2) {
													XSSFCell Rcell_obj = rowobj.getCell(j,rowobj.RETURN_BLANK_AS_NULL);
																	
													if (Rcell_obj != null) {
														cell_val5 = Rcell_obj.getStringCellValue();
														alName.add(cell_val5);
													}
												}

												// Loop through all the fieldvalue
												// columns and store the values in all
												// fieldvalues in the
												// a1Val array list
												String cell_val6 = null;
												for (int j = 6; j <= Cell_count; j = j + 2) {
													XSSFCell Rcell_obj = rowobj.getCell(j,rowobj.RETURN_BLANK_AS_NULL);
															
													if (Rcell_obj != null) {
														cell_val6 = Rcell_obj.getStringCellValue();
														alVal.add(cell_val6);		
													}
												}

												// if atleast one fieldname and
												// fieldvalue are present
												if (alName.size() > 0 && alVal.size() > 0) {
													GenericFunctions.validateOperation(alName, alVal);
												} else {
													Logger.warn("No input data is available for validation!!");
												}
												// gets the status and the result string
												// and reports it using the reportresult
												// function
												int resultst = GenericFunctions.exe_rst_status;
												String rsmssg = GenericFunctions.returnresultmssg();
														
												reportResult(resultst,"The expected value"+ cell_val6+ " should matched with the actual"+ cell_val5 + " value",
														rsmssg);
											}
										}
										
										
										
										
										int resultst = GenericFunctions.exe_rst_status;
										String rsmssg = GenericFunctions.returnresultmssg();
										reportResult(resultst,"The expected value should present",rsmssg);
												
										
									}
									// this block is executed if the keyword is
									// input
									else if (Keyword.equalsIgnoreCase("INPUT")) {

										// creates two array list of type string
										// one for object name and second for
										// object values
										ArrayList<String> alName = new ArrayList<String>();
										ArrayList<String> alVal = new ArrayList<String>();

										// stores the values for object names in
										// a1Name array list
										String cell_val5 = null;
										int Cell_count = rowobj
												.getLastCellNum();
										for (int j = 5; j <= Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null) {
												cell_val5 = Rcell_obj
														.getStringCellValue();
												alName.add(cell_val5);
											}
										}

										// stores the value for object values in
										// a1val array list
										String cell_val6 = null;
										for (int j = 6; j <= Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null) {
												cell_val6 = Rcell_obj
														.getStringCellValue();
												alVal.add(cell_val6);
											}
										}

										// This function is called to do
										// different types of input operation
										// like
										// entering value in a text field,
										// selecting a value from a listbox or
										// combobox
										// selecting/deselecting a radio button
										// or checbox etc
										GenericFunctions.INPUT(alName, alVal);
										// get the final result of the input
										// operation
										int resultst = GenericFunctions.exe_rst_status;
										// report the result for the input
										// operation
										reportResult(
												resultst,
												"The value "
														+ cell_val6
														+ " Should be entered in the "
														+ cell_val5 + " field",
												"The value "
														+ cell_val6
														+ " is entered successfully in the "
														+ cell_val5 + " field",
												"The value "
														+ cell_val5
														+ " could not be entered in the "
														+ cell_val6 + " field");
									}

									// this block is executed if the keyword is
									// click
									else if (Keyword.equalsIgnoreCase("Click")) {

										// it gets the fieldname and fieldvalue
										XSSFCell Cell_obj5 = rowobj.getCell(5);
										XSSFCell Cell_obj6 = rowobj.getCell(6);

										if (Cell_obj5 != null) {
											String ObjectName = Cell_obj5
													.getStringCellValue();
											OR_ObjectName = ObjectName;
											// create a new array list arr
											ArrayList<String> arr = new ArrayList<String>();
											// if object name has multiple
											// values seperated using comma this
											// functions returns
											// the elements seperately in an
											// arraylist
											arr = GenericFunctions
													.returnArraylistStringCommaSeprated(OR_ObjectName);
											// System.out.println("page object present in excel : "+arr.size());
											Logger.info("page object present in excel : "
													+ arr.size());
											// if there are more than 2 elements
											// present in the array
											if (arr.size() >= 2
													&& Cell_obj6 != null) {
												// System.out.println("Under multiple object handle operation!!");
												Logger.info("Under multiple object handle operation!!");
												// if clicking on multiple
												// objects in one operation and
												// value of fieldvalue is
												// also present e.g textbased
												GenericFunctions
														.handleDynamicOperation(
																"click",
																Cell_obj6
																		.toString());
											}
											// if only element present i.e. we
											// dont have multiple values
											// seperated by comma
											else if (arr.size() == 1) {
												GenericFunctions
														.clickCondition(
																OR_ObjectName,
																Cell_obj6);
											}
										} else {
											// other operation for click like
											// clicing in a link
											GenericFunctions
													.OtherClickOperation(Cell_obj6);
										}
										// get the final status and report the
										// result
										int resultst = GenericFunctions.exe_rst_status;
										reportResult(
												resultst,
												OR_ObjectName
														+ " Object should be clicked",
												OR_ObjectName
														+ " Object is clicked",
												OR_ObjectName
														+ " Object is not clicked");

									}

									// if keyword is wait
									else if (Keyword.equalsIgnoreCase("wait")) {
										XSSFCell Cell_obj3 = rowobj.getCell(5);
										// get the value of the object from
										// fieldname1
										String ObjectName = Cell_obj3
												.getStringCellValue();

										OR_ObjectName = ObjectName;

										// call the waitoperation and pass the
										// fieldname and fieldvalue as parameter
										GenericFunctions.waitOperation(
												rowobj.getCell(5),
												rowobj.getCell(6));
										// the return value of the waitoperation
										int resultst = GenericFunctions.exe_rst_status;
										String resutmsg = GenericFunctions.resultmssg;
										// report the result
										reportResult(resultst,
												"Application should wait!!",
												resutmsg);
									}
									// if keyword is waitforTheElement
									else if (Keyword
											.equalsIgnoreCase("waitForTheElement")) {
										XSSFCell Cell_obj3 = rowobj.getCell(5);
										String ObjectName = Cell_obj3
												.getStringCellValue();

										OR_ObjectName = ObjectName;

										// pass the object name as parameter

										GenericFunctions.waitFortheElement(500,
												ObjectName);
										int resultst = GenericFunctions.exe_rst_status;
										String resutmsg = GenericFunctions.resultmssg;
										reportResult(resultst,
												"Application should wait!!",
												resutmsg);
									}

									// if keyword is close
									else if (Keyword.equalsIgnoreCase("close")) {
										// function to close single of multiple
										// windoes
										GenericFunctions.Close();
										int resultst = GenericFunctions.exe_rst_status;
										reportResult(resultst,
												"Window Should be Closed",
												"Window closed successfully",
												"Window not closed");
									}

									// block to be executed if keyword is
									// ValidateExcelData
									else if (Keyword.equalsIgnoreCase("ValidateExcelData")) {

										ArrayList<String> FieldName = new ArrayList<String>();
										ArrayList<String> FieldValue = new ArrayList<String>();
										// creates two arraylist one for
										// fieldnames and second for fieldvalues
										int Total_Cell_count = rowobj
												.getLastCellNum();
										String cell_value5 = null;
										for (int j = 5; j <= Total_Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null)
											{
												objFormulaEvaluator.evaluate(Rcell_obj);
												cell_value5 = objDefaultFormat.formatCellValue(Rcell_obj,objFormulaEvaluator);
												//cell_value5 = Rcell_obj
													//	.getStringCellValue();
												FieldName.add(cell_value5);
											}
										}

										String cell_value6 = null;
										for (int j = 6; j <= Total_Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null) 
											{
												//objFormulaEvaluator.evaluate(Rcell_obj);
												Rcell_obj.setCellType(Cell.CELL_TYPE_STRING);
												//cell_value6 = objDefaultFormat.formatCellValue(Rcell_obj,objFormulaEvaluator);
												cell_value6 = Rcell_obj
													.getStringCellValue();
												FieldValue.add(cell_value6);
											}
										}
										// calls the function to validate the
										// data in an excel file and passes the
										// excel sheet path, fieldname and
										// fieldvalue array lists as parameter
										GenericFunctions.validate_Excel_Data(
												Temp_ResultSheetPath,
												FieldName, FieldValue);

										// get the result from the function,
										// report result
										int resultst = GenericFunctions.exe_rst_status;
										String resutmsg = GenericFunctions.resultmssg;
										reportResult(resultst,
												"Excel file validated!!",
												resutmsg);
									}

									else if ((Keyword
											.equalsIgnoreCase("GetValue")))

									{

										ArrayList<String> GV_FieldName = new ArrayList<String>();
										ArrayList<String> GV_FieldValue = new ArrayList<String>();
										// creates two arraylist one for
										// fieldnames and second for fieldvalues
										int Total_Cell_count = rowobj
												.getLastCellNum();
										String cell_value5 = null;
										for (int j = 5; j <= Total_Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null) {
												cell_value5 = Rcell_obj
														.getStringCellValue();
												GV_FieldName.add(cell_value5);
											}
										}

										String cell_value6 = null;
										for (int j = 6; j <= Total_Cell_count; j = j + 2) {
											XSSFCell Rcell_obj = rowobj
													.getCell(
															j,
															rowobj.RETURN_BLANK_AS_NULL);
											if (Rcell_obj != null) {
												cell_value6 = Rcell_obj
														.getStringCellValue();
												GV_FieldValue.add(cell_value6);
											}
										}
										if (GV_FieldName.size() > 0
												&& GV_FieldValue.size() > 0) {

											GenericFunctions
													.GetValue(GV_FieldName,
															GV_FieldValue);
										} else {
											Logger.warn("No input data is available for fetching value!!");
										}

										// gets the status and the result string
										// and reports it using the reportresult
										// function
										int resultst = GenericFunctions.exe_rst_status;
										String rsmssg = GenericFunctions
												.returnresultmssg();
										reportResult(
												resultst,
												"The value stored in "
														+ cell_value5
														+ " should be fetched and stored",
												rsmssg);

									}

									// this block is executed if keyword is
									// other and is used for operations like
									// executing
									// vbscript, autoit, keypress, switching to
									// iframe etc
									else if (Keyword.equalsIgnoreCase("other")) {
										String action = GenericFunctions
												.fn_Data("action");
										if (action != null) {
											String validtedata = action;
											char[] data = validtedata
													.toCharArray();
											if (data.length > 0) {
												GenericFunctions
														.executeOtherOperation(validtedata);
											}
										}
										int resultst = GenericFunctions.exe_rst_status;
										String resutmsg = GenericFunctions.resultmssg;
										reportResult(
												resultst,
												action
														+ " Operation should be success !!",
												resutmsg);
									}

									// This block is executed if keyword is
									// openURL
									else if (Keyword.equalsIgnoreCase("OpenURL")) {
										String Browser;
										XSSFCell Cell_obj6 = rowobj.getCell(6);
										String URL = Cell_obj6.getStringCellValue();
										try {
											Browser = Driver.browsername;
											
										} catch (Exception e) {
											Browser = "";
											Browser = "FF";
										}

										URL = URL.trim();
										// calls the openurl function and
										// supplies url and browser as parameter
										
										GenericFunctions.OpenURL(URL, Browser);
										// gets the final status and reports the
										// result
										int resultst = GenericFunctions.exe_rst_status;
										reportResult(resultst, " " + URL
												+ " should be launched", ""
												+ URL
												+ " is launched successfully",
												"" + URL + " is not launched");
									}
									// this block is executed if keyword is
									// FileUpload

									else if (Keyword
											.equalsIgnoreCase("FileUpload"))

									{
										XSSFCell Cell_obj3 = rowobj.getCell(5);
										String ObjectName = Cell_obj3
												.getStringCellValue();
										XSSFCell Cell_obj5 = rowobj.getCell(6);
										String Path = "";
										String Path1 = "";
										String Path2 = "";

										if (Cell_obj5 != null) {
											Path1 = Resourse_path.currPrjDirpath
													+ "\\test-data\\";
											Path2 = Cell_obj5
													.getStringCellValue();
											Path = Path1 + Path2;
										}

										OR_ObjectName = ObjectName;
										// function to upload file, passes the
										// filename and filepath as parameter
										GenericFunctions.fileUpload(
												OR_ObjectName, Path);
										int resultst = GenericFunctions.exe_rst_status;
										reportResult(
												resultst,
												" " + OR_ObjectName
														+ " should be uploaded",
												""
														+ OR_ObjectName
														+ " is upload successfully",
												""
														+ OR_ObjectName
														+ " is not uploaded successfully");
									}
									// clearing the find test data for current
									// row
									Bl.clear();
								}

							}
						}
					} catch (Exception e) {
						Logger.warn("exceptionHandler called!!");
						String messagedata = GenericFunctions.getStackTrace(e);
						exceptionreport("Object should be found",
								"Object not found!!", messagedata);
						Logger.error(e);
						e.printStackTrace();
					} finally {
						if (SkipTestCase != "True") {
							if (CurrentTestID.equalsIgnoreCase(nextTestID)) {
								if (TestStatus.equalsIgnoreCase("failed"))
									continue;
							} else if (((CurrentTestID != nextTestID) || (nextTestID == null))
									&& (k <= Row_count)) {
								Logger.info("Row " + CurrentTestID + "!="
										+ nextTestID + " || (" + nextTestID
										+ "==null)) && (" + k + "<="
										+ Row_count + ")");
								if (TestStatus.equalsIgnoreCase("Passed")
										|| TestStatus
												.equalsIgnoreCase("Failed")) {
									Logger.info("Status :: " + TestStatus
											+ " :: Finalizing result report");
									PrevTestID = null;
									Date dte = new Date();
									DateFormat df = DateFormat
											.getTimeInstance();
									String E_Time = df.format(dte);
									End_Time = E_Time;
									try {
										int Length_status = Arr_list.size();
										for (int m = 0; m < Length_status; m++) {
											String Status = (String) Arr_list
													.get(m);
											// System.out.println("Status "+Status);
											if (Status
													.equalsIgnoreCase("Failed")) {
												TestStatus = "Failed";
												break;
											} else {
												TestStatus = "Passed";
											}
										}
										String[] Array = { Driver.S_No,
												Driver.Testcasename,
												Driver.start_time,
												Driver.End_Time,
												Driver.TestStatus };
										GenericFunctions
												.Write_HL_ResultToExcel(Array,
														Driver.HL_RSheetName);
										Arr_list.clear();
										GenericFunctions
												.DrawGraph(Driver.HL_RSheetName);

										if (CurrentTestID != nextTestID) {
											TestStatus = "passed";
										}

									} catch (Exception e) {
										e.printStackTrace();
										// System.out.println("Exception in driver class:: "+
										// e.getMessage());
										Logger.error(e);
									}

									if (TestStatus.equalsIgnoreCase("Failed")) {
										PrevTestID = null;
										TestStatus = "passed";
									}
								}

							}

							if (k == Row_count) {
								if (GenericFunctions.driver != null) {
									GenericFunctions.driver.quit();
									GenericFunctions.driver = null;
									Logger.info("Driver is quit");
									Logger.info("Execution for "
											+ DriverSheetname
											+ " is completed..Now executing next application..");
									Logger.info("\n");
									Logger.info("/*************************************************************************/");
									Logger.info("/**************   Execution completed           **************************/");
									Logger.info("/**************************************************************************/");
									Logger.info("\n");
								}
								// Assigning null values to driver on complete
								// execution

								GenericFunctions.driver = null;
							}

							Wbook_obj.close();
						}
					}
				}
				// Need to add code here for send email
			}
		}

		// force quit the IEDriver process thread
		// GenericFunctions.killProcess("IEDriverServer");
		// Force quit to java execution
		System.exit(-1);
	}

	// add step report to excel
	public static void reportResult(int resultst, String Expectedmsg,
			String resultmssg) {

		try {
			if (GenericFunctions.winCount() >= 1) {
				if (GenericFunctions.driver != null) {
					if (Snap_flag.equalsIgnoreCase("always")) {
						Snap_URL = GenericFunctions.Fn_TakeSnapShotAndRetPath(GenericFunctions.driver);
					}
				}
			}
			Logger.info("restult status : " + resultst);
			if (resultst == 1) {
				String[] Resultarray = { Driver.Testcasename,
						Driver.StepNumber, Driver.HL_RSheetName,
						Driver.KeywordName, Driver.ScreenName, Expectedmsg,
						resultmssg, "Passed", Snap_URL };
				GenericFunctions.WriteResultToExcel(Resultarray,Driver.HL_RSheetName);
				Arr_list.add("Passed");
				TestStatus = "Passed";
			}

			else if (resultst == 2) {
				String[] Resultarray = { Driver.Testcasename,
						Driver.StepNumber, Driver.HL_RSheetName,
						Driver.KeywordName, Driver.ScreenName, Expectedmsg,
						resultmssg, "Failed", Snap_URL };
				GenericFunctions.WriteResultToExcel(Resultarray,
						Driver.HL_RSheetName);
				Arr_list.add("Failed");
				TestStatus = "Failed";
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// add step report to excel
	public static void reportResult(int resultst, String Expectedmsg,
			String resultpssmssg, String resultfailmssg) {
		try {
			if (GenericFunctions.winCount() >= 1) {
				if (GenericFunctions.driver != null) {
					if (Snap_flag.equalsIgnoreCase("always")) {
						// This captures the screenshot and returns the full
						// path of the png file
						Snap_URL = GenericFunctions
								.Fn_TakeSnapShotAndRetPath(GenericFunctions.driver);
					}
				}
			}
			Logger.info("restult status : " + resultst);
			// if pass
			if (resultst == 1) {
				// creates the entire result string i.e Test case name, Test
				// step number, Application name, Keyword
				// screen name, the expected result, actual results, Passed and
				// snapshot url
				String[] Resultarray = { Driver.Testcasename,
						Driver.StepNumber, Driver.HL_RSheetName,
						Driver.KeywordName, Driver.ScreenName, Expectedmsg,
						resultpssmssg, "Passed", Snap_URL };
				// this function writes the results to the detailed report
				GenericFunctions.WriteResultToExcel(Resultarray,
						Driver.HL_RSheetName);
				// stores the result passed om Arr_List and also in the
				// TestStatus variable
				Arr_list.add("Passed");
				TestStatus = "Passed";

			}
			// If status is failed then
			else if (resultst == 2) {
				// create the result array
				String[] Resultarray = { Driver.Testcasename,
						Driver.StepNumber, Driver.HL_RSheetName,
						Driver.KeywordName, Driver.ScreenName, Expectedmsg,
						resultfailmssg, "Failed", Snap_URL };
				// call the function to write the results to the detailed report
				// sheet
				GenericFunctions.WriteResultToExcel(Resultarray,
						Driver.HL_RSheetName);
				// store the result Failed in Arr_List and also set the value of
				// Test Status as Failed
				Arr_list.add("Failed");
				TestStatus = "Failed";
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// add step report to excel
	public static void exceptionreport(String Expectedmsg, String resultmssg,
			String excmsg) {
		try {
			if (GenericFunctions.winCount() >= 1) {
				if (GenericFunctions.driver != null) {
					if (Snap_flag.equalsIgnoreCase("always")) {
						Snap_URL = GenericFunctions
								.Fn_TakeSnapShotAndRetPath(GenericFunctions.driver);
					}
				}
			}
			String[] Resultarray = { Driver.Testcasename, Driver.StepNumber,
					Driver.HL_RSheetName, Driver.KeywordName,
					Driver.ScreenName, Expectedmsg, resultmssg, "Failed",
					Snap_URL, excmsg };
			GenericFunctions.WriteResultToExcel(Resultarray,
					Driver.HL_RSheetName);
			Arr_list.add("Failed");
			TestStatus = "Failed";
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
