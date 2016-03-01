package configuration;

import java.io.File;

import genericclasses.GenericFunctions;
import driver.Driver;

public class Resourse_path {
	
	public static String currPrjDirpath = System.getProperty("user.dir"); //To get current directory path
	
	public static String DateStamp = GenericFunctions.fn_GetCurrentDate();
	
	public static String selenium_version = "2.45.0";
	
	public static String IEDriver_version = "2.45.0";
	
	public static String DateTimeStamp = GenericFunctions.fn_GetCurrentTimeStamp();
	
	public static String applicationresultfolder = Driver.HL_RSheetName;
	
	public static String ie_32bit = "32bit_2.45.0"+File.separator+"IEDriverServer_64bit_2.45.0.exe";
	
	public static String ie_64bit = "64bit"+File.separator+"IEDriverServer_64bit_2.45.0.exe";
	
	public static String chrome_driver_path = currPrjDirpath+"/resources/browser-drivers/Chrome Driver/chromedriver.exe";
		
	public static String testcaseresultfolder = Driver.Testcasename;
		
	public static String homepath = System.getProperty("user.home"); //To get home path
	
	public static String Driver_Sheetpath = currPrjDirpath+"/Schedular/Suite Driver.xlsx"; //To get Driver Sheet path
	
	public static String Driver_SheetName = "Driver"; // To get Driver Sheet tab name
	
	public static String TestData_Sheetpath = currPrjDirpath+"/TestCases/"; //To get test data sheet path
	
	public static String browserDriverpath = currPrjDirpath+"/resources/browser-drivers/"+ie_64bit;
	
	public static String autoItpath = currPrjDirpath+"/resources/autoIt/compiled/";
	
	public static String vbfolderpath = currPrjDirpath+"/resources/vbscript/";
	
	public static String browsertoolpath = currPrjDirpath+"/resources/tool/";
	
	public static String Result_Sheet_header[]= {"TestcaseName","TestStepNumber" , "ApplicationName", "TestStep" , "Screen", "ExpectedResult" , "ActualResult" , "Status" , "SnapshotLink"}; // Detailed result sheet columns

	public static String HighLevel_Result_header[] = {"SR.No" , "TestCaseName", "Start Time" , "End Time", "Status"}; // High level result sheet column
	
	public static String date_time_down_folder = "";
	
	public static String comp_downloadfoder = "";
	
	public static String logfilepath = "";
	
	public static String running_path = "";
	
	public static String tempfolder = homepath+File.separator+"AppData"+File.separator+"Local"+ File.separator+"Temp"+File.separator;
	
	public static String prj_download_fld = currPrjDirpath+File.separator+"Downloads"+File.separator;
	
	public static String flag_a = ""; // It stores the value till complete execution
	
	public static String sys_def_down_fold_path = homepath+File.separator+"Downloads"+File.separator; // It stores defualt system download path
	
	public static String csv_path = currPrjDirpath+"/test-data/";

}
