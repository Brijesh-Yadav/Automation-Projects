package genericclasses;



import java.awt.AWTException;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.commons.io.FileUtils;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerDriverService;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import configuration.Resourse_path;
import driver.Driver;

public class GenericFunctions {
	
	//Declared statics object and variables
	public static WebDriver driver;
	public static String Snapshot_flg;
	public static String Browser_Type;
	public static String EmployeeCsvFname;
	public static String DependentCsvFname;
	public static String HistoryCsvFname;
	public static ArrayList<String> Snp_flag= new ArrayList<String>();
	public static ArrayList<String> empfname= new ArrayList<String>();
	public static ArrayList<String> depfname= new ArrayList<String>();
	public static ArrayList<String> hisfname= new ArrayList<String>();
	public static String mainwindow;
	public static String resultmssg = null;
	public static int exe_rst_status = 0; // 0 for no report, 1 for pass and 2 for fail
	
	
	public static String returnresultmssg(){
		return resultmssg;
	}
	
	//This function iterates through the driver sheet or scheduler sheet file and checks which have to be executed using the Exec indicator value
	//If the indicator is Yes then adds the sheet name to A1 Array List. It also picks up the snapshot indicator and stores in Snp_flag array list
	//need to trap specific errors like file not found, no rows in file, or no file has execution indicator as yes
	
	public static ArrayList<String> ReadDriverSuiteExcel() {
		ArrayList<String> Al = new ArrayList<String>();
		XSSFWorkbook Workbook_obj = null;
		try {
			String Driversheetpath = Resourse_path.Driver_Sheetpath;
			FileInputStream FIS = new FileInputStream(Driversheetpath);
			Workbook_obj = new XSSFWorkbook(FIS);
			XSSFSheet sheet_obj = Workbook_obj.getSheet("Driver");
			int Row_count = sheet_obj.getLastRowNum();
			 System.out.print(Row_count+"\n");
			for (int i = 0; i <= Row_count; i++) {
				XSSFRow row_obj = sheet_obj.getRow(i);
				XSSFCell cell_obj = row_obj.getCell(3);
				
				if(cell_obj!=null){
					String Exec_indicator = cell_obj.getStringCellValue();
//					 System.out.print(Exec_indicator+"\n");
					String Exec_ind = Exec_indicator.trim();
					if ("Y".equalsIgnoreCase(Exec_ind)     ) {
						XSSFCell cellobj1 = row_obj.getCell(1);
						String Sheetname = cellobj1.getStringCellValue();
						Al.add(Sheetname);
						
						XSSFCell cellobj4 = row_obj.getCell(4);
						Snapshot_flg = cellobj4.getStringCellValue();
						Snp_flag.add(Snapshot_flg);
						
						XSSFCell cellobj5 = row_obj.getCell(5);
						Browser_Type = cellobj5.getStringCellValue();
						/*
						//get the value of the employee cvs file name
						XSSFCell cellobj6 = row_obj.getCell(6);
						if (row_obj.getCell(6)!=null)
						{
						String EmployeeCsvname = cellobj6.getStringCellValue();
						empfname.add(EmployeeCsvname);
						}
						else
						{
							empfname.add("NA");
						}
						//get the value of the dependent csvfile name
						XSSFCell cellobj7 = row_obj.getCell(7);
						if (row_obj.getCell(7)!=null)
						{
						String DependentCsvname = cellobj7.getStringCellValue();
						depfname.add(DependentCsvname);
						}
						else
						{
							depfname.add("NA");
						}
						
						//get the value of the history csv file name
						XSSFCell cellobj8 = row_obj.getCell(8);
						if (row_obj.getCell(8)!=null)
						{
							String HistoryCsvname = cellobj8.getStringCellValue();
							hisfname.add(HistoryCsvname);
						}
						else
						{
							hisfname.add("NA");
						}
						*/
				}
				
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}finally{
			try {
				Workbook_obj.close();
			} catch (IOException e) {
				e.printStackTrace();
				Logger.error(e);
			}
		}
		return Al;
	}
	
	
	//This function iterates through the driver sheet or scheduler sheet file and checks which have to be executed using the Exec indicator value
		//If the indicator is Yes then adds the sheet name to A1 Array List. It also picks up the snapshot indicator and stores in Snp_flag array list
		//need to trap specific errors like file not found, no rows in file, or no file has execution indicator as yes
		
		public static ArrayList<String> ReadTestCaseSchedulerExcel(String SheetPath) {
			ArrayList<String> TC_All = new ArrayList<String>();
			XSSFWorkbook Workbook_obj = null;
			try {
				FileInputStream FIS = new FileInputStream(SheetPath);
				Workbook_obj = new XSSFWorkbook(FIS);
				XSSFSheet sheet_obj = Workbook_obj.getSheet("TestCaseScheduler");
				int Row_count = sheet_obj.getLastRowNum();
				// System.out.print(Row_count+"\n");
				for (int i = 1; i <= Row_count; i++) {
					XSSFRow row_obj = sheet_obj.getRow(i);
					XSSFCell cell_obj = row_obj.getCell(2);
					String Exec_indicator = cell_obj.getStringCellValue();
					// System.out.print(Exec_indicator+"\n");
					String Exec_ind = Exec_indicator.trim();
					if (Exec_ind.equalsIgnoreCase("yes")) {
						XSSFCell cellobj1 = row_obj.getCell(0);
						String TCName = cellobj1.getStringCellValue();
						TC_All.add(TCName);
											
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
				Logger.error(e);
			}finally{
				try {
					Workbook_obj.close();
				} catch (IOException e) {
					e.printStackTrace();
					Logger.error(e);
				}
			}
			return TC_All;
		}
		
		
//This function is to get all the parameters i.e. fieldname and fieldvalue in a arraylist and return the arraylist		
		
	public static ArrayList<String> FindTestData() {
		
		//name of the test case sheet to be executed
		String TestData_Sheetpath = Resourse_path.TestData_Sheetpath+ Driver.DriverSheetname + ".xlsx";
        //create a new array list D1
		ArrayList<String> Dl = new ArrayList<String>();
		//create a new workbook object
		XSSFWorkbook Wbook_obj = null;
		//store values of stepnumber, testcaseid, keyword and test case sheet in variables
		String stp = Driver.StepNumber;
		String script_Tcid = Driver.Testcasename;
		String keyword = Driver.KeywordName;
		String TestData_SheetName = Driver.SuiteName;
		
		try {
			//open the test case excel file
			FileInputStream FIS = new FileInputStream(TestData_Sheetpath);
			Wbook_obj = new XSSFWorkbook(FIS);
			//get access to the test case sheet and store reference in a variable
			XSSFSheet Wsheet_obj = Wbook_obj.getSheet("Test Case");
			//get the total count in the excel sheet test case
			int Rowcount = Wsheet_obj.getLastRowNum();
			int RequiredRow = 0;
//Loop through all the rows in the excel sheet test case
			for (int i = 1; i <= Rowcount; i++) {
				//creating row object
				XSSFRow rowobj = Wsheet_obj.getRow(i);
				//Storing values to string variable for stepnum, test case id, keyword
				XSSFCell Cell_obj = rowobj.getCell(1);
				String Excel_stepnum = Cell_obj.getStringCellValue();
				XSSFCell Cell_obj1 = rowobj.getCell(0);
				String Excel_Tcid = Cell_obj1.getStringCellValue();
				XSSFCell Cell_obj2 = rowobj.getCell(4);
				String Excel_Keyword = Cell_obj2.toString();
				//match the values of the current row an stored the rownum in the required row variable
				if (stp.equalsIgnoreCase(Excel_stepnum) && Excel_Keyword.equalsIgnoreCase(keyword) && 
						Excel_Tcid.equalsIgnoreCase(script_Tcid)) {
					RequiredRow = i;
					break;
				}
				
			}
            //get the row object using the required row pointer
			XSSFRow RRow_obj = Wsheet_obj.getRow(RequiredRow);
			//get the total number of cells in the current row
			int Cell_count = RRow_obj.getLastCellNum();
			//start looping through all the columns starting column number 5
			for (int j = 5; j <= Cell_count - 1; j++) {
				//get the cell object
				XSSFCell Rcell_obj = RRow_obj.getCell(j);
				if(Rcell_obj!=null){
					//get the value in a variable
					String cell_val = Rcell_obj.getStringCellValue();
					//add the value in the D1 arraylist 
					Dl.add(cell_val);
				}
			}
			
		} 
		//catch block for catching exception and printing stacktrace and also logging error in log file
		catch (Exception e) {
            e.printStackTrace();
            Logger.error(e);
		}
		//finally block to close the wbook 
		finally {
			try {
				Wbook_obj.close();
			  } catch (IOException e) {
				e.printStackTrace();
				Logger.error(e);
			}	
		}
		//returning the array list
		return Dl;
	}
	
	public static void CreateHighLevelResult(String Result_SheetName) {
		try {
			String Result_Sheetpath = Driver.HighLevel_Result_Sheetpath;
			String[] Result_Sheet_header = Resourse_path.HighLevel_Result_header;
            //Excel operation
			XSSFWorkbook Wbook_obj = new XSSFWorkbook();
			XSSFSheet Wsheet_obj = Wbook_obj.createSheet(Result_SheetName);
			XSSFRow Row_obj = Wsheet_obj.createRow(0);
			int Col_loop = Result_Sheet_header.length;
			for (int i = 0; i < Col_loop; i++) {
				XSSFCell cell_obj = Row_obj.createCell(i);
				cell_obj.setCellValue(Result_Sheet_header[i]);
				CellStyle cellStyleobj = SetCellstyl(Wbook_obj);
				cell_obj.setCellStyle(cellStyleobj);
				Wsheet_obj.autoSizeColumn(i);
			}
			FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
			Wbook_obj.write(FOS);
			FOS.close();
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	
	public static void Format_HL_Result(String arr[], String Result_SheetName) {
		try {
			String Result_Sheetpath = Driver.HighLevel_Result_Sheetpath;

			FileInputStream FIS = new FileInputStream(Result_Sheetpath);
			Workbook Wbook_obj1 = WorkbookFactory.create(FIS);
			
			if (Wbook_obj1.getSheet(Result_SheetName) == null) {
				String[] Result_Sheet_header = Resourse_path.HighLevel_Result_header;
				FileInputStream N_FIS = new FileInputStream(Result_Sheetpath);
				Workbook Wbook_obj = WorkbookFactory.create(N_FIS);
				Sheet Wsheet = Wbook_obj.createSheet(Result_SheetName);
				
				Row Row_obj = Wsheet.createRow(0);
				int Col_loop = Result_Sheet_header.length;
				
				for (int i = 0; i < Col_loop; i++) {
					Cell cell_obj = Row_obj.createCell(i);
					cell_obj.setCellValue(Result_Sheet_header[i]);
					CellStyle cellStyleobj = SetCellstyl(Wbook_obj);
					cell_obj.setCellStyle(cellStyleobj);
					Wsheet.autoSizeColumn(i);
				}
				FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
				Wbook_obj.write(FOS);
				FOS.close();

			}
			FIS.close();
			
			//Second operation
			FileInputStream FIS_1 = new FileInputStream(Result_Sheetpath);
			Workbook Wbook_obj = WorkbookFactory.create(FIS_1);
			Sheet Wsheet_obj = Wbook_obj.getSheet(Result_SheetName);
			
			int Rowcount = Wsheet_obj.getLastRowNum();
			int ReqRow_Num = Rowcount + 1;
			Row NSobj = Wsheet_obj.createRow(ReqRow_Num);
			
			for (int i = 0; i < arr.length; i++) {
				// String Value = null;
				String Value = arr[i];
				
				Cell CellObj = NSobj.createCell(i);

				if (i != 4) {
					CellObj.setCellValue(Value);
					CellStyle Borderstyle_obj = GenericFunctions.SetCellBorderstyl(Wbook_obj);
					CellObj.setCellStyle(Borderstyle_obj);
				}
				else if (i == 4) {
					if (Value == "Passed") {
						CellStyle CellStyleObj = Wbook_obj.createCellStyle();
						Font Fontobj = Wbook_obj.createFont();
						Fontobj.setFontHeightInPoints((short) 11);
						Fontobj.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
						Fontobj.setColor(IndexedColors.GREEN.getIndex());
						CellStyleObj.setFont(Fontobj);
						CellStyleObj.setBorderLeft(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderRight(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderTop(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderBottom(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setLeftBorderColor(IndexedColors.BLACK.getIndex());
						CellStyleObj.setRightBorderColor(IndexedColors.BLACK.getIndex());
						CellStyleObj.setTopBorderColor(IndexedColors.BLACK	.getIndex());
						CellStyleObj.setBottomBorderColor(IndexedColors.BLACK.getIndex());
						CellObj.setCellStyle(CellStyleObj);
						CellObj.setCellValue(Value);
					} 
					else if (Value == "Failed") {
						CellStyle CellStyleObj = Wbook_obj.createCellStyle();
						Font Fontobj = Wbook_obj.createFont();
						Fontobj.setFontHeightInPoints((short) 11);
						Fontobj.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
						Fontobj.setColor(IndexedColors.RED.getIndex());
						CellStyleObj.setFont(Fontobj);
						CellStyleObj.setBorderLeft(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderRight(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderTop(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setBorderBottom(XSSFCellStyle.BORDER_THIN);
						CellStyleObj.setLeftBorderColor(IndexedColors.BLACK.getIndex());
						CellStyleObj.setRightBorderColor(IndexedColors.BLACK.getIndex());
						CellStyleObj.setTopBorderColor(IndexedColors.BLACK.getIndex());
						CellStyleObj.setBottomBorderColor(IndexedColors.BLACK.getIndex());
						CellObj.setCellStyle(CellStyleObj);
						CellObj.setCellValue(Value);
					}
				}
			}
			
			FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
			Wbook_obj.write(FOS);
			FOS.close();
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	}

	public static void Write_HL_ResultToExcel(String Res_Arr [], String Sheetname) {
		GenericFunctions.Format_HL_Result(Res_Arr,Sheetname+"_High_Level");
	}
	
	public static void CreateResultExcel(String Result_SheetName){
		try{
			String Result_Sheetpath = Driver.LowLeveL_Result_Sheetpath;
			// String Result_SheetName=Resourse_path.Result_SheetName;
			String[] Result_Sheet_header = Resourse_path.Result_Sheet_header;

			XSSFWorkbook Wbook_obj = new XSSFWorkbook();
			XSSFSheet Wsheet_obj = Wbook_obj.createSheet(Result_SheetName);
			XSSFRow Row_obj = Wsheet_obj.createRow(0);
			int Col_loop = Result_Sheet_header.length;
			for (int i = 0; i < Col_loop; i++) {
				XSSFCell cell_obj = Row_obj.createCell(i);
				cell_obj.setCellValue(Result_Sheet_header[i]);
				CellStyle cellStyleobj = SetCellstyl(Wbook_obj);
				cell_obj.setCellStyle(cellStyleobj);
				Wsheet_obj.autoSizeColumn(i);
			}
			File files = new File(Driver.LowLevel_Result_Folder);
			if (!files.exists()) {
				files.mkdirs();
			}
			FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
			Wbook_obj.write(FOS);
			FOS.close();
			
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	
 //************************** Code for font ***********************
	public static Font Fontstyle(Workbook wbookobj, short fontheight) {
		Font Fontobj = wbookobj.createFont();
		Fontobj.setFontHeightInPoints(fontheight);
		Fontobj.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		Fontobj.setColor(IndexedColors.WHITE.getIndex());
		return Fontobj;
	}
	
	//********************* Code for Cell style ***************
	public static CellStyle SetCellBorderstyl(Workbook wbookobj) {
		CellStyle cellBorderStyleobj = wbookobj.createCellStyle();
		cellBorderStyleobj.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellBorderStyleobj.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellBorderStyleobj.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellBorderStyleobj.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellBorderStyleobj.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		cellBorderStyleobj.setRightBorderColor(IndexedColors.BLACK.getIndex());
		cellBorderStyleobj.setTopBorderColor(IndexedColors.BLACK.getIndex());
		cellBorderStyleobj.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		return cellBorderStyleobj;
	}

	public static CellStyle SetCellstyl(Workbook wbookobj) {
		CellStyle cellStyleobj = wbookobj.createCellStyle();
		short colouInd = IndexedColors.BLACK.getIndex();
		short fillpatern = CellStyle.SOLID_FOREGROUND;
		short fontheight = 11;
		cellStyleobj.setFillForegroundColor(colouInd);
		cellStyleobj.setFillPattern(fillpatern);
		Font Fontobj = Fontstyle(wbookobj, fontheight);
		cellStyleobj.setFont(Fontobj);
		return cellStyleobj;
	}

	public static CellStyle SetCellstyl_LLResult(Workbook wbookobj) {
		CellStyle cellStyleobj = wbookobj.createCellStyle();
		short colouInd = IndexedColors.YELLOW.getIndex();
		short fillpatern = CellStyle.SOLID_FOREGROUND;
		short fontheight = 11;
		cellStyleobj.setFillForegroundColor(colouInd);
		cellStyleobj.setFillPattern(fillpatern);
		Font Fontobj1 = Fontstyle_LLResult(wbookobj, fontheight);
		cellStyleobj.setFont(Fontobj1);
		return cellStyleobj;
	}

	public static Font Fontstyle_LLResult(Workbook wbookobj, short fontheight) {
		Font Fontobj1 = wbookobj.createFont();
		Fontobj1.setFontHeightInPoints(fontheight);
		Fontobj1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		Fontobj1.setColor(IndexedColors.BLACK.getIndex());
		return Fontobj1;
	}
	
	public static void AddRowToResult(String Result_SheetName) {
		try {
			String Result_Sheetpath = Driver.LowLeveL_Result_Sheetpath;
			Logger.info("Result_Sheetpath" + Result_Sheetpath);
			FileInputStream FIS = new FileInputStream(Result_Sheetpath);
			Workbook Wbook_obj = WorkbookFactory.create(FIS);
			Sheet Wsheet_obj = Wbook_obj.getSheet(Result_SheetName+ "_Detailed_Results");
			int LastRownum = Wsheet_obj.getLastRowNum();
			int requiredrow = LastRownum + 1;
			Row Row_obj = Wsheet_obj.createRow(requiredrow);
			for (int i = 0; i < 9; i++) {
				Cell cell_obj = Row_obj.createCell(i);
				CellStyle cellStyleobj = SetCellstyl_LLResult(Wbook_obj);
				cell_obj.setCellStyle(cellStyleobj);
				Wsheet_obj.autoSizeColumn(i);
			}
			// Cell_obj.setCellValue("END");
			FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
			Wbook_obj.write(FOS);
			FOS.close();
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//This function writes the results to the detailed report and also formats it
	public static void FormatResultExcel(String arr [],String Result_SheetName) throws IOException, InvalidFormatException{
		
		String Result_Sheetpath = Driver.LowLeveL_Result_Sheetpath;
		//gets the file input stream, workbook and worksheet object
		FileInputStream FIS_1 = new FileInputStream(Result_Sheetpath);
		Workbook Wbook_obj = WorkbookFactory.create(FIS_1);
		Sheet Wsheet_obj = Wbook_obj.getSheet(Result_SheetName);
		//gets the total number of rows in the worksheet
		int Rowcount = Wsheet_obj.getLastRowNum();
		int ReqRow_Num = Rowcount + 1;
		Row NSobj = Wsheet_obj.createRow(ReqRow_Num);
		int lencmp = arr.length;
//		Logger.info("lencmp "+lencmp);
		String exmsg = null;
		if(lencmp==10){
			exmsg = arr[9];
		}
		for (int i = 0; i < arr.length; i++) {
			String Value = arr[i];
			Cell CellObj = NSobj.createCell(i);
			if (i != 7 && i != 8 && i!=9) {
				CellObj.setCellValue(Value);
				CellStyle Borderstyle_obj = GenericFunctions.SetCellBorderstyl(Wbook_obj);
				CellObj.setCellStyle(Borderstyle_obj);
			} else if (i == 8) {
				// Logger.info("i==8");
				if(Value!=null){
					if(Snapshot_flg.equalsIgnoreCase("Always")){
						CellStyle hlink_style = Wbook_obj.createCellStyle();
						Font hlink_font = Wbook_obj.createFont();
						hlink_font.setUnderline(Font.U_SINGLE);
						hlink_font.setColor(IndexedColors.BLUE.getIndex());
						hlink_style.setFont(hlink_font);
						CellObj.setCellValue("Link");
						CreationHelper createHelper = Wbook_obj.getCreationHelper();
						Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_FILE);
						Value = Value.replace("\\", "/");
						link.setAddress(Value);
						CellObj.setHyperlink((org.apache.poi.ss.usermodel.Hyperlink) link);
						CellObj.setCellStyle(hlink_style);
					}
					
				}
				
			} else if (i == 7) {
				if (Value == "Passed") {
					CellStyle CellStyleObj = Wbook_obj.createCellStyle();
					Font Fontobj = Wbook_obj.createFont();
					Fontobj.setFontHeightInPoints((short) 11);
					Fontobj.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
					Fontobj.setColor(IndexedColors.GREEN.getIndex());
					CellStyleObj.setFont(Fontobj);
					CellStyleObj.setBorderLeft(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderRight(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderTop(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderBottom(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setLeftBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setRightBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setTopBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setBottomBorderColor(IndexedColors.BLACK
							.getIndex());

					CellObj.setCellStyle(CellStyleObj);
					CellObj.setCellValue(Value);
				} else if (Value == "Failed") {

					CellStyle CellStyleObj = Wbook_obj.createCellStyle();
					Font Fontobj = Wbook_obj.createFont();
					Fontobj.setFontHeightInPoints((short) 11);
					Fontobj.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
					Fontobj.setColor(IndexedColors.RED.getIndex());
					CellStyleObj.setFont(Fontobj);
					CellStyleObj.setBorderLeft(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderRight(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderTop(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setBorderBottom(XSSFCellStyle.BORDER_THIN);
					CellStyleObj.setLeftBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setRightBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setTopBorderColor(IndexedColors.BLACK
							.getIndex());
					CellStyleObj.setBottomBorderColor(IndexedColors.BLACK
							.getIndex());
					CellObj.setCellStyle(CellStyleObj);
					CellObj.setCellValue(Value);
					
					if(lencmp==10){
						if(exmsg!=null){
							fillcommentinexcel(Wbook_obj,Wsheet_obj,NSobj,CellObj,exmsg);
						}else{
							Logger.info("exmsg msg "+exmsg);
						}
					}
					
				}
			}
			Wsheet_obj.autoSizeColumn(i);
		}
		FileOutputStream FOS = new FileOutputStream(Result_Sheetpath);
		Wbook_obj.write(FOS);
		FOS.close();
	}
	
	//This function creates the result sheet if it does not exist and also calls
	//the formarresultexcel function and passes the result array to it
	public static void WriteResultToExcel(String Res_Arr [],String Sheetname){
		try{
			//get the path to the detailed sheet
			String Result_Sheetpath=Driver.LowLeveL_Result_Sheetpath;
			//get the file object
			File ResultFile_obj=new File(Result_Sheetpath);
			//if file does not exist then create the file
			if (ResultFile_obj.exists()==false){
				GenericFunctions.CreateResultExcel(Sheetname+"_Detailed_Results");
			}//format the excel file
			GenericFunctions.FormatResultExcel(Res_Arr,Sheetname+"_Detailed_Results");
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	
	public static String fn_Data(String arg_FieldName) {
		String xl_FiledValue = null;
	try
	{
		//Array List with all the data from the test case sheet
		ArrayList<String> TD_AL = Driver.FL;
		//stores the number of elements in the array list
		int dataCnt = TD_AL.size();
		//String xl_FiledValue = null;
		//Loops through all the elements in the arraylist
		for (int i = 0; i <= dataCnt-1; i++) {
			//gets the Fieldname
			String xl_FiledName = (String) TD_AL.get(i);
			if (xl_FiledName.equalsIgnoreCase(arg_FieldName)) {
				xl_FiledValue = (String) TD_AL.get(i + 1);
				Logger.info("FiledName "+xl_FiledName+" and FiledValue "+xl_FiledValue);
				break;
			}
		}
	}
		catch(Exception e)
		{
			Logger.warn("activate window : fail :: error occurs :: using different function to capture image");
		}
		
		return xl_FiledValue;
	
	
	};
		
		
	// *********Function for Current Date
	public static String fn_GetCurrentDate() {
		Date dte = new Date();
		DateFormat df = DateFormat.getDateInstance();
		String strdte = df.format(dte);
		strdte = strdte.replaceAll(":", "_");
		strdte = strdte.replaceAll(",", "_");
		strdte = strdte.replaceAll(" ", "_");
		return strdte;
	}

	// *********Function for Current Date
	public static String fn_GetCurrentTimeStamp() {
		Date dte = new Date();
		DateFormat df = DateFormat.getDateTimeInstance();
		String strdte = df.format(dte);
		strdte = strdte.replaceAll(":", "_");
		strdte = strdte.replaceAll(",", "_");
		strdte = strdte.replaceAll(" ", "_");
		return strdte;
	}
	 
	  public static String Fn_TakeSnapShotAndRetPath(WebDriver WebDriver_Object){
		  String FullSnapShotFilePath = null ;
		  String ex_snapshotPath = null ;
		  
		  try{
			  /*
			  //window focus
			  if(Driver.browsername.equalsIgnoreCase("ff")){
				if (winCount() == 1) {
					// driverting focus to default window
					Set<String> windows_chk = driver.getWindowHandles();
					for (String window : windows_chk) {
						driver.switchTo().window(window);
					}
					Logger.info("Window focused for screenshot!!");
				}
			  }else if(Driver.browsername.equalsIgnoreCase("ie")){
				String activate = fn_Data("defaultwindow");
				if (activate != null) {
					if (activate.equalsIgnoreCase("activate")) {
						// driverting focus to window
						Set<String> windows_chk = driver.getWindowHandles();
						for (String window : windows_chk) {
							driver.switchTo().window(window);
						}
					}
				}
			  }
			  */
			  //get the current time
			  String TimeStamp = GenericFunctions.fn_GetCurrentTimeStamp();
			  //get the current folder for storing snapshots 
			  String FolderPath = Driver.Temp_ResultSheetPath+"/Snapshots";
			  
			  //String FolderPath = Resourse_path.currPrjDirpath +"/Results"+"/Results_"+Resourse_path.DateTimeStamp+"/Snapshots";
			  //create a file object
			  File FolderObj=new File(FolderPath);
			  //create the snapshots folder
			  FolderObj.mkdir();
			  //get the absolute path for the snapshots folder
			  FolderPath=FolderObj.getAbsolutePath();
			  //get the name of the screen name and replace all spaces with underscore and store in a string variable
			  String Snap_ScreenName=Driver.ScreenName.replaceAll(" ", "_");
			  //create the full path for the snashot file by combining
			  //snapshot folder path and test case sheet name, test case name, screen name 
			  //and keyword and timespamp.png
			  String snapshot_path = Driver.HL_RSheetName +"/"+Driver.Testcasename+"/"+Snap_ScreenName+"__"+Driver.KeywordName+"__("+TimeStamp+").png";
			  FullSnapShotFilePath=FolderPath+"/"+snapshot_path;
			  ex_snapshotPath = "Snapshots/"+snapshot_path;
			  
			  //create reference to this new file 
			  File DestFile=new File(FullSnapShotFilePath);
			  //take screenshot using webdriver object
			  //
			  File SrcFile=((TakesScreenshot)WebDriver_Object).getScreenshotAs(OutputType.FILE);
			  //copy the file to the destfile object
			  FileUtils.copyFile(SrcFile, DestFile);
			  
			  
		  }
		  //catch block for screenshot code
		  //catch the exception and log and also try and take screenshot using java code
		  catch(Exception e){
			  Logger.warn("Take Screenshot : fail :: error occurs :: using different function to capture image");
			  captureScreen(FullSnapShotFilePath);
		  }
		  //return the full path of the screenshot file
		  return ex_snapshotPath;
	   }
    
	//Take screenshot using java function  
	public static void captureScreen(String fileName) {
		try {
			Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
			Rectangle screenRectangle = new Rectangle(screenSize);
			Robot robot = new Robot();
			BufferedImage image = robot.createScreenCapture(screenRectangle);
			ImageIO.write(image, "png", new File(fileName));
			Logger.info("Screenshot captured using java function!!");
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	} 
	 // The title of the window is passed as parameter to this function and it loops through all
	// the windows and matches the title provided as parameter and the title of the current window
	//if window is switched successfully sets flag as true otherwise calls function to switch window based
	//on object
	public static void handleMultipleWindows(String windowTitle) {
		Logger.info("Title in xml file - "+windowTitle);
		//get all the window handles in the array windows
		Set<String> windows = driver.getWindowHandles();
		//this is flag to be set if window is found, default value set to false
		boolean window_ana = false;
		//Looping through the string array windows
		for (String window : windows) {
			//switching to open windows one by one
			driver.switchTo().window(window);
			//if the title passed as parameter matches with the title of the current window
			if (driver.getTitle().equalsIgnoreCase(windowTitle)) {
				//set flag to true
				window_ana = true;
				Logger.info("Title of the page after - switching window To: "+ driver.getTitle());
				Logger.info("Desired window is activated!!");
				//maximize the window
				driver.manage().window().maximize();
				break;
			}
		}
		// if window not found then log warning and call the function to handle window bases on object
		if (window_ana != true) {
			Logger.warn("Window is not switched, now trying based on element!!");
			handleMultipleWindowsObjectBased();
		}
	}
	
	public static void handleMultipleWindowsObjectBased() {
		
		// gets the handle of all open windows in a array
		Set<String> windows = driver.getWindowHandles();
		//sets flag to false
		boolean window_ana = false;
		
		//loop through all windows one by one
		for (String window : windows) {
			//switch to window
			driver.switchTo().window(window);
			Logger.info("Window title "+driver.getTitle());
			//calls the function that returns the webelement object
			List<WebElement> fieldobjs = Field_objs(Driver.OR_ObjectName);
			//sets the flag for windows found to true and maximizes the windows and comes out
			if(fieldobjs.size()>0){
				Logger.info("Desired window is activated!!");
				window_ana = true;
				driver.manage().window().maximize();
				break;
			}
		}
		// switching result-
		if (window_ana != true) {
			Logger.warn("Window is not switched !!");
		}
	}
	
//this function returns the number of open windows
	  
	public static int winCount() throws InterruptedException {
		int winCount = 0;
		try{
			if (driver != null) {
				Thread.sleep(1000);
				//get the count of all open windows
				winCount = driver.getWindowHandles().size();
				Logger.info("Current WebDriver Win Count is:: " + winCount);
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		return winCount;
	}
	  
	public static void DrawGraph(String HL_Result_SheetName) {

		try{
			int pass_count = 0;
			int Fail_count = 0;

			FileInputStream chart_file_input = new FileInputStream(new File(Driver.HighLevel_Result_Sheetpath));
			XSSFWorkbook my_workbook = new XSSFWorkbook(chart_file_input);
			XSSFSheet my_sheet = my_workbook.getSheet(HL_Result_SheetName	+ "_High_Level");
			DefaultPieDataset my_pie_chart_data = new DefaultPieDataset();
			
			int rownum = my_sheet.getLastRowNum();
			for (int i = 1; i <= rownum; i++) {
				XSSFRow my_Row = my_sheet.getRow(i);
				XSSFCell my_cell = my_Row.getCell(4);
				String Status = my_cell.getStringCellValue();
				if (Status.equalsIgnoreCase("Passed")) {
					pass_count = pass_count + 1;
				} else if (Status.equalsIgnoreCase("Failed")) {
					Fail_count = Fail_count + 1;
				}
			}
			String Status_p = "Passed";
			String Status_f = "Failed";

			my_pie_chart_data.setValue(Status_p, pass_count);
			my_pie_chart_data.setValue(Status_f, Fail_count);

			JFreeChart myPieChart = ChartFactory.createPieChart(
					"Execution Status", my_pie_chart_data, true, true, false);
			
			PiePlot plot = (PiePlot) myPieChart.getPlot();
			plot.setSectionPaint(Status_f, Color.RED);
			plot.setSectionPaint(Status_p, Color.GREEN);
			plot.setBackgroundPaint(Color.white);
			
			plot.setLabelGenerator(new StandardPieSectionLabelGenerator(
					"{0}{2}", NumberFormat.getNumberInstance(), NumberFormat.getPercentInstance()));
			int width = 500;
			int height = 500;
			float quality = 1;
			
			ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
			ChartUtilities.writeChartAsJPEG(chart_out, quality, myPieChart,width, height);
			
			InputStream feed_chart_to_excel = new ByteArrayInputStream(chart_out.toByteArray());
			
			byte[] bytes = IOUtils.toByteArray(feed_chart_to_excel);
			int my_picture_id = my_workbook.addPicture(bytes,	Workbook.PICTURE_TYPE_JPEG);
			
			feed_chart_to_excel.close();
			chart_out.close();
			XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
			ClientAnchor my_anchor = new XSSFClientAnchor();
			my_anchor.setCol1(7);
			my_anchor.setRow1(2);
			XSSFPicture my_picture = drawing.createPicture(my_anchor,my_picture_id);
			my_picture.resize();
			chart_file_input.close();
			FileOutputStream out = new FileOutputStream(new File(Driver.HighLevel_Result_Sheetpath));
			my_workbook.write(out);
			out.close();
			my_workbook.close();
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//This is the function that accepts username and password and inputs the username and password and clicks
	//on the Loginbutton and also returns the status as 1 i.e. Passed
	public static void Login(String UNAME, String PASSWORD) {
		exe_rst_status = 2;
		//username
		executeInputOperation("UserName", UNAME);
		//password
		executeInputOperation("Password", PASSWORD);
		//button
		Click("LoginButton");
		exe_rst_status = 1;
	}
	
	/*public static void INPUT(String OR_ObjectName, String Value){
		String fl = OR_ObjectName.trim();
		if (fl.contains("robot_")) {
			executeRobotevent(fl);
		} else {
			executeInputOperation(OR_ObjectName, Value);
		}
	}*/
	
	//This function accepts two arraylists as parameters one for object names and second for object values
	public static void INPUT(ArrayList<String> alName, ArrayList<String> alVal){
		int sA = alName.size();
			int sB = alVal.size();
			//if both array list have same number of elements
			if (sA==sB) {
				
				//loop through the object name array list
				for (int i = 0; i < sA; i++) {
						
					//get the name of the object and name of the value in variables
						String oName = alName.get(i);
						String oVal = alVal.get(i);
						
						executeInputOperation(oName, oVal);

				}
			}
		}
	
	//this function accepts the object name and the value to enter in the text field as parameter
	public static void EnterText(WebElement Object_name, String ValueToEnter) {
		Logger.info("ValueToEnter "+ValueToEnter);
		//waitFortheElement(500, Object_name);
		//clear the text filed
		Object_name.clear();
		//enter the value in the textfield using sendkeys
		Object_name.sendKeys(ValueToEnter);
		Logger.info("Entered value is "+ValueToEnter);
		//check if the value is correctly entered
		boolean vluecheck1 = validateTextfieldvalue(Object_name,ValueToEnter);
		//if correct value is not entered
		if(vluecheck1!=true){
			//Filling value again if the value is not entered
			Logger.info("Try - Enter data with action builder");
			//craete an object for the actions class
			Actions builder = new Actions(driver);
			//enter value using actions class
			builder.moveToElement(Object_name).sendKeys(ValueToEnter).build().perform();
			//check if correct value is entered
			boolean vluecheck2 = validateTextfieldvalue(Object_name,ValueToEnter);
			//if correct value is not entered using actions class also
			if(vluecheck2!=true){
				Logger.info("Try - Enter data with robot action");
				//enter the text using robot class
				typecharacter(ValueToEnter);
				staticwait(1);
				//check if the correct value is entered using robot class
				boolean vluecheck3 = validateTextfieldvalue(Object_name,ValueToEnter);
				//validating entered value
				//if correct value is not entered using robot class also
				if(vluecheck3!=true){
					Logger.info("Try - Enter data with autoit");
					//get the exepath of generic functions.exe
					String exefilepath = Resourse_path.autoItpath+"genericfunctions"+".exe";
					try {
						@SuppressWarnings("unused")
						//enter the value using autoit
						Process pb= new ProcessBuilder(exefilepath,"EnterValue",ValueToEnter).start();
					} catch (IOException e) {
						e.printStackTrace();
					}
					//check if the correct value is entered using autoit
					boolean vluecheck4 = validateTextfieldvalue(Object_name,ValueToEnter);
					//if value is still not entered correctly
					if(vluecheck4!=true){
						Logger.info("Value is not entered in text box");
						//check if object is disabled if yes do logging
						if (Object_name.isEnabled()==false){
							Logger.info("The required textbox is not enabled therefore "
									+ ValueToEnter + " cannot be entered");
						}
					}
				}
			}
		}
	}
	
	public static void typecharacter(String s){
		try{
			//create an object for the robot class
			Robot robot = new Robot();
			byte[] bytes = s.getBytes();
			for (byte b : bytes) {
				int code = b;
				// keycode only handles [A-Z] (which is ASCII decimal [65-90])
				if (code > 96 && code < 123)
					code = code - 32;
				robot.delay(40);
				robot.keyPress(code);
				robot.keyRelease(code);
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//this function accepts the webelement and the value for the webelement
	public static boolean validateTextfieldvalue(WebElement ele, String vl){
		exe_rst_status = 2;
		boolean rsult =false;
		//gets the value for the webelement and stores in a variable
		String attvalue = ele.getAttribute("value");
		Logger.info("Attribute entered value "+attvalue);
//		Logger.info("Attribute entered value "+attvalue);
		//if value supplied as parameter and the current value for the webelement matches then set status as passed
		//and return the result as true otherwise false
		
		if(attvalue.equalsIgnoreCase(vl)){
			rsult = true;
			exe_rst_status = 1;
		}
		return rsult ; 
	}
	
	//this function accepts the object name and value incase of checkbox value will be select or unselect
	public static void SelectCheckbox(WebElement Object_name,String flag){
		 exe_rst_status = 2;
//	     WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
		 //get the current status of the checkbox i.e true or false
	     boolean currentFlag = Object_name.isSelected();
	    //if operation is to select and the state is already selected then set status as passed
	     if (flag.equalsIgnoreCase("select") || flag==null || flag.equals("")){
	    	 if (currentFlag==true){
//	    		 Logger.info("The checkbox  "+Object_name+" is already selected");
	    		 Logger.info("The checkbox  "+Object_name+" is already selected");
	    		 exe_rst_status = 1;
	    	 }
	    	 
	    	 //select the checkbox, set status as passed
	    	 else
	    	 {
	    		 Object_name.click();
//	    		 Logger.info("Checkbox is selected!!");
	    		 Logger.info("Checkbox is selected!!");
	    		 exe_rst_status = 1;
	    	 }
	     }
	     //if operation is to unselect 
	     else if(flag.equalsIgnoreCase("unselect")) 
	     {
	    	 // if currently the chechbox is selected then click and unselect and set status as passed
	    	 if (currentFlag==true) {
	    		 Object_name.click();
//	    		 Logger.info("Checkbox is unselected!!");
	    		 Logger.info("Checkbox is unselected!!");
	    		 exe_rst_status = 1;
	    	 }
	    	 //if currently status is unselected then set status as passed and do log entry
	    	 else{
//	    		 Logger.info("The checkbox  "+Object_name+" is already De-Selected");
	    		 Logger.info("The checkbox  "+Object_name+" is already De-Selected");
	    		 exe_rst_status = 1;
	    	 }
	     } 
	}
	
	//this function accepts the object for the radio button and the value select or unselect
	public static void RadioSelector(WebElement Object_name,String flag){
		 exe_rst_status = 2;
//	     WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
		 //get the current status of the radio button
	     boolean currentFlag = Object_name.isSelected();
	    
	     //if operation is to select and the radio button is already selected then set status as passed
	     if (flag.equalsIgnoreCase("select") || flag==null || flag.equals("") ){
	    	 if (currentFlag==true){
//	    		 Logger.info("The Radio Button  "+Object_name+" is already selected");
	    		 Logger.info("The Radio Button  "+Object_name+" is already selected");
	    		 exe_rst_status = 1;
	    	 }
	    	 
	    	 //if current status is not already selected then select the radio button
	    	 else
	    	 {
	    		 Object_name.click();
	    		 exe_rst_status = 1;
	    	 }
	     }
	     
	     //if operation is to unselect and current status is selected then click the radio button and unselect
	     else if (flag.equalsIgnoreCase("unselect")){
	    	 if (currentFlag==true) {
	    		 Object_name.click();
	    		 exe_rst_status = 1;
	    	 }
	    	 //else if it is already unselected then set status as passed and do logging
	    	 else {
//	    		 Logger.info("The Radio Button  "+Object_name+" is already De-Selected");
	    		 Logger.info("The Radio Button  "+Object_name+" is already De-Selected");
	    		 exe_rst_status = 1;
	    	 }
	     } 
	     //executed if there is someother value then select or unselect
	     else{
	    	 Logger.warn("select is not found in excel");
	     }
	}
	
	public static void SelectFromListBox(WebElement Object_name,String ValueToSelect) {
		exe_rst_status = 2;	
		//if sheet has action as fieldname it gets the fieldvalue in action variable
		String action = fn_Data("action");
		
		// if action has a value
		if(action!=null)
		{
			char [] lngh = action.toCharArray();
			if(lngh.length>0){
				//if action has value javascript_click
				if(action.equalsIgnoreCase("javascript_click")){
					Logger.info("javascript click");
					JavascriptExecutor js = null;
		            if (driver instanceof JavascriptExecutor) {
		                js = (JavascriptExecutor)driver;
		            }
		            //select value from listbox using javascript
					js.executeScript("arguments[0].click()", Object_name);
					exe_rst_status = 1;	
				}
				//if sheet has value for action fieldname as sendkeys
				else if(action.equalsIgnoreCase("sendkeys")){
					Logger.info("sendkeys");
					//Object_name.sendKeys("abcd");
					//gets the value from the excel sheet
					String e_value = fn_Data("value");
					//creates a list of webelement for options tag
					List<WebElement> optionsize = Object_name.findElements(By.tagName("option"));
					//gets the total count of the optionsize list
					int opsize = optionsize.size();
					Logger.info("optionsize "+opsize);
					//loop through the options list
					for(int j=1; j<=opsize; j++){
						try {
							//creates an object of robot class
							Robot robot = new Robot();
							//gets the value attribute
							String ac_vl = Object_name.getAttribute("value");
							Logger.info("dp_vl "+ac_vl);
							//if value in excel and application match, slect the value using keyrelease and keypress
							if(e_value.equalsIgnoreCase(ac_vl)){
								robot.keyPress(KeyEvent.VK_TAB);
								robot.keyRelease(KeyEvent.VK_TAB);
								break;
							}else{
								robot.keyPress(KeyEvent.VK_DOWN);
								robot.keyRelease(KeyEvent.VK_DOWN);
							}
						} catch (AWTException e) {
							e.printStackTrace();
							Logger.error(e);
						}
					}
					exe_rst_status = 1;
				}
			}
		}
		
		//if there is no special action
		else 
		{
			//create the object for the object name passed as parameter
			Select Object_NM = new Select(Object_name);
			String value ="";
			try{
				//get the value selected
				value = Object_NM.getFirstSelectedOption().getText();
				//if value to be selected is already passed as parameter then status as passed and log entry
				if(value.equalsIgnoreCase(ValueToSelect)){
					Logger.info(ValueToSelect+" Value is already selected in dropdown!!");
					exe_rst_status = 1;
				}
				else
				{
					if (!Object_NM.isMultiple()){
						//waitFortheElement(500, ValueToSelect);
						//select the value in the list box and status as passed and log entry
						Object_NM.selectByVisibleText(ValueToSelect);
						Logger.info(ValueToSelect+" value selected!!");
						exe_rst_status = 1;
					}
					//incase of a multiselect listbox set status as failed and do log entry 
					else
					{
						Logger.info(Object_NM+ " is multiselect drop down list.. Use different Action for this object");
						exe_rst_status = 2;	
					}
				}
			}
			//catch block for function
			catch(Exception e){
				Logger.warn("Exception handled for dropdown function");
				//calling the select method again if it is not a multiselect list box
				if (!Object_NM.isMultiple()){
					Object_NM.selectByVisibleText(ValueToSelect);
					Logger.info(ValueToSelect+" value selected!!");
					exe_rst_status = 1;
				}
				//else for multiselect list box
				else
				{
					Logger.info(Object_NM+ " is multiselect drop down list.. Use different Action for this object");
					exe_rst_status = 2;	
				}
			}
			
		}
	}
	
	//this function accepts the object name to be clicked as parameter and clicks on it
	public static void Click(String OR_ObjectName) {
		exe_rst_status = 2;
		Logger.info("Checking for object in click");
		//returns the webelement for the object to be clicked
		WebElement Object_name = GenericFunctions.Field_obj(OR_ObjectName);
		//if browser is ie and we want to use sendkeys or javascript click
		if (Driver.browsername.equalsIgnoreCase("ie")) {
			String operation = fn_Data("operation");
			if(operation!=null){
				if(operation.equalsIgnoreCase("sendkeys")){
					waitFortheElement(120, OR_ObjectName);
					Object_name.sendKeys(Keys.ENTER);
					Logger.info("Sendkeys click performed on " + OR_ObjectName);
				}
				else if(operation.equalsIgnoreCase("javascript_click")){
					waitFortheElement(120, OR_ObjectName);
					JavascriptExecutor js = (JavascriptExecutor) driver;
					js.executeScript("arguments[0].click()", Object_name);
					Logger.info("javascript click perfomed on "+OR_ObjectName);
				}
			}
			//normal click operation
			else
			{
				
				String TC_Name = Driver.TestCaseID;
				 if(TC_Name.equalsIgnoreCase("TC1"))
			 
				{
					 waitFortheElement(620, OR_ObjectName);
					 Logger.info("clicking on " + OR_ObjectName);
					 Object_name.click();
					 driver.manage().timeouts().pageLoadTimeout(220, TimeUnit.SECONDS);
					 //((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
					 Logger.info("clicked on " + OR_ObjectName);
				}
				 else
				 {
					 waitFortheElement(120, OR_ObjectName);
					 Logger.info("clicking on " + OR_ObjectName);
					 Object_name.click();
					 driver.manage().timeouts().pageLoadTimeout(220, TimeUnit.SECONDS);
					 //((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
					 Logger.info("clicked on " + OR_ObjectName);
				 }
				
			}
			exe_rst_status = 1;
		} 
		
		//if browser is ff and normal click
		else if (Driver.browsername.equalsIgnoreCase("ff"))
		{
			String TC_Name = Driver.TestCaseID;
			 if(TC_Name.equalsIgnoreCase("TC1"))
		 
			{
				 waitFortheElement(620, OR_ObjectName);
				 Object_name.click();
					Logger.info("clicked on " + OR_ObjectName);
					exe_rst_status = 1;
			}
			 else
			 {
				 waitFortheElement(120, OR_ObjectName);
				 Object_name.click();
					Logger.info("clicked on " + OR_ObjectName);
					exe_rst_status = 1;
			 }
			
		} //if browser is ff and normal click
		else if (Driver.browsername.equalsIgnoreCase("chrome"))
		{
			String TC_Name = Driver.TestCaseID;
			 if(TC_Name.equalsIgnoreCase("TC1"))
		 
			{
				 waitFortheElement(620, OR_ObjectName);
				 Object_name.click();
					Logger.info("clicked on " + OR_ObjectName);
					exe_rst_status = 1;
			}
			 else
			 {
				 waitFortheElement(120, OR_ObjectName);
				 Object_name.click();
					Logger.info("clicked on " + OR_ObjectName);
					exe_rst_status = 1;
			 }
			
		} 
		else {
			Logger.info("Browser name is not mentioned for last open url!!");
		}
	}
	
	public static boolean handleMultipleWindowsforClosingwindow() {
		boolean rs = false;
		String br_window = fn_Data("browser_window");
		
		if(br_window!=null){
			Logger.info("br_window "+br_window);
			boolean window_ana = false;
			if(br_window.equalsIgnoreCase("title")){
				Set<String> windows = driver.getWindowHandles();
				for (String window : windows) {
					driver.switchTo().window(window);
					String obj_title = driver.getTitle();
					Logger.info("Window title "+driver.getTitle());
					String obj_value = fn_Data("obj_value");
					if(obj_value!=null){
						char [] val_len = obj_value.toCharArray();
						if(val_len.length>0){
							if(obj_title.equalsIgnoreCase(obj_value)){
								Logger.info("Desired window is activated!!");
								window_ana = true;
								rs = true;
								break;
							}
						}
					}else{
						Logger.info("obj_value value null in excel");
					}
				}
			} else if(br_window.equalsIgnoreCase("index")){
				Set<String> windows = driver.getWindowHandles();
				String obj_value = fn_Data("obj_value");
				if(obj_value!=null){
					char [] val_len = obj_value.toCharArray();
					if(val_len.length>0){
						if(Integer.parseInt(obj_value)>=windows.size()){
							//not implemented yet
							//need to write logic here to handle window on index based
							rs = true;
						}
					}
				}else{
					Logger.info("obj_value value null in excel");
				}
				
			} else if(br_window.equalsIgnoreCase("url")){
				Set<String> windows = driver.getWindowHandles();
				for (String window : windows) {
					driver.switchTo().window(window);
					String page_url = driver.getCurrentUrl();
					Logger.info("Window url "+page_url);
					String obj_value = fn_Data("obj_value");
					if(obj_value!=null){
						char [] val_len = obj_value.toCharArray();
						if(val_len.length>0){
							if(page_url.equalsIgnoreCase(obj_value)){
								Logger.info("Desired window is activated!!");
								window_ana = true;
								rs = true;
								break;
							}
						}
					}else{
						Logger.info("obj_value value null in excel");
					}
				}
			}else {
				Set<String> windows = driver.getWindowHandles();
				for (String window : windows) {
					driver.switchTo().window(window);
					Logger.info("Window switched to - "+driver.getTitle());
					List<WebElement> fieldobjs = Field_objs(br_window);
					if(fieldobjs.size()>0){
						Logger.info("Desired window is activated!!");
						window_ana = true;
						rs = true;
						break;
					}
				}
			}
			
			// switching result-
			if (window_ana != true) {
				Logger.info("Window is not switched !!");
			}
			
		}
		return rs ;
	}
	
	public static void Close() {
		exe_rst_status = 2;
		//passes the windowApp_name as parameter to fn_Data function this function accepts fieldname as 
		//parameter and returns the fieldvalue
		String processName = fn_Data("windowApp_name");

		if (processName != null && !processName.isEmpty()) {
		//	killProcess(processName);
		} else {
			Logger.info("closing window...");
			
			// This passes closewindow as parameter to fn_Data and gets the fieldvalue for closeWindow fieldname
			//from the excel sheet as the return value
			String closewindow = fn_Data("closewindow");
			//if close window has some value
			if (closewindow != null)
			{
				//if parameter for closewindow is closeall then do driver.quit and set status as passed
				if (closewindow.equalsIgnoreCase("closeall")) {
					driver.quit();
					driver=null;
					// return result status
					exe_rst_status = 1;
				}
				
			}
			
			//if fieldname is closewindow but there is no fieldvalue parameter
			else
			{
				
				//get number of open windows
				int wincnt = driver.getWindowHandles().size();
				Logger.info("window size " + wincnt);
				
				// if one wondow is open
				if (wincnt == 1) {
					driver.manage().deleteAllCookies();
					//((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
					//close the window
					driver.quit();
					
					driver = null;
					Logger.info("Window is closed!!");
					exe_rst_status = 1;
				} 
				//if windows open are more than one
				else 
				{
					//this switches windos if more than one window is open 
					boolean vl = handleMultipleWindowsforClosingwindow();
					//if switch is unsuccessfull, try switching again
					if (vl != true) {
						// activating top window
						Set<String> wind_prs = driver.getWindowHandles();
						for (String window : wind_prs) {
							driver.switchTo().window(window);
							String url = driver.getCurrentUrl();
							String title = driver.getTitle();
							Logger.info("url " + url + " and title is " + title);
						}
					}

					// close browser instance
					((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
					driver.close();
					Logger.info("Window is closed!!");
					int win = driver.getWindowHandles().size();
					if(win==1){
						Set<String> windows = driver.getWindowHandles();
						for (String window : windows) {
							driver.switchTo().window(window);
							break;
						}
					}else if(win>=2){
						Set<String> windows = driver.getWindowHandles();
						for (String window : windows) {
							driver.switchTo().window(window);
						}
					}
					
					String closew = fn_Data("closewindow");
					if(closew!=null){
						if(closew.equalsIgnoreCase("checkagain")){
							// checking widow presence again
							int wincnt2 = driver.getWindowHandles().size();
							if (wincnt2 == 1) {
								Set<String> windows = driver.getWindowHandles();
								// activating top window
								for (String window : windows) {
									driver.switchTo().window(window);
									String url = driver.getCurrentUrl();
									String title = driver.getTitle();
									Logger.info("url " + url + " and title is " + title);
								}
								// checking window session
								// close browser instance
								((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
								driver.close();
								Logger.info("Window is closed!!");
							} else if (wincnt2 >= 2) {
								Set<String> windows = driver.getWindowHandles();
								// activating top window
								for (String window : windows) {
									driver.switchTo().window(window);
									String url = driver.getCurrentUrl();
									String title = driver.getTitle();
									Logger.info("url " + url + " and title is " + title);
									Set<String> win_sz = driver.getWindowHandles();
									if (win_sz.size() == 1) {
										break;
									}
									// browser will not get closed where there is only 1
									// window
									// checking window session
									// close browser instance
									((JavascriptExecutor)driver).executeScript("isUnSubmittedForms = false;");
									driver.close();
									Logger.info("Window is closed!!");
								}
							}
						}
					}
					// return result status
					exe_rst_status = 1;
				}
			}
		}
	}
	
	public static void checkwindowsession(){
		// checking active window session
		String checksession2 = driver.getWindowHandle();
		if (checksession2 != null) {
			driver.switchTo().defaultContent();
			Logger.info("session is not null :: " + checksession2);
		}
	}
	
	public static void quitDriver(){
		Set<String> windows = driver.getWindowHandles();
        Logger.info("window size "+windows.size());
        int windowsize = windows.size();
        if (windowsize == 1) {
            String session = driver.getWindowHandle();
            if(session !=null){
                 driver.quit();
            }else{
                 //activating top window
                 for (String window : windows) {
                      driver.switchTo().window(window);
                      driver.quit();
                      Logger.info("window closed and driver is quit");
                 }
            }
		} else {
			//activating top window
			for (String window : windows) {
				driver.switchTo().window(window);
				String url = driver.getCurrentUrl();
				String title = driver.getTitle();
				Logger.info("url "+url+" and title is "+title);
				driver.quit();
				staticwait(1);
				Logger.info("window closed and driver is quit");
			}
		}
		// quiting the driver instance
		if (driver != null) {
			driver.quit();
		}
		// Assigning null values to driver
		driver = null;
	}
	
	//
	public static void Validate(String OR_ObjectName, String Exp_val){
		exe_rst_status = 2;
		resultmssg = null;
		//This finds the object and returns the webelemnt object
		WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
		int getvalcheck = Exp_val.indexOf("|");
		String Expval = Exp_val;
		if (getvalcheck>= 0)
		{
			Logger.info("Key value is " + Exp_val);
			Expval = Driver.GVmap.get(Exp_val);
		}
		
		Logger.info("Exp_value "+Expval);
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		//get the actual value from the object using getText
		String ActValue = Object_name.getText();
		//converts the value to lowercase
		Logger.info("Actual Value"+ ActValue);
		//compares the actual and expected values, sets status and result message string accordingly
		if (ActValue.equalsIgnoreCase(Expval)){
			exe_rst_status = 1;
			resultmssg = "Exp : "+Expval+" and act : "+ActValue+" value matched!!";
		} else {
			resultmssg = "Exp : "+Expval+" and act : "+ActValue+" value not matched!!";
			exe_rst_status = 2;
		}	
	}
	
	
	public static void isElementPresent(String OR_ObjectName) {
		WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
		if (Object_name.isDisplayed() == true) {
	        System.out.println("My element was found on the page");
	} else
	        System.out.println("My element was not found on the page");
	}
	
	
	

	public static void Validate(String OR_ObjectName,String operator,String VariableType,String Value){
		
		int ReqValue=0;
		int ActValue=0;
		Boolean Flag=false;
		exe_rst_status = 2;
		resultmssg = null ;
		
		WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
		String ActualValue = Object_name.getText();
		Logger.info("Actual value "+ActualValue);
		
		if(VariableType.equalsIgnoreCase("Number")){
			
			ReqValue=Integer.parseInt(Value);
			ActValue=Integer.parseInt(ActualValue);
			
			if(operator.equalsIgnoreCase("Equal to")){
				if (ActValue==ReqValue) {
					Flag=true;
				}
				
			} else if(operator.equalsIgnoreCase("Not Equal to")){
				
				if (ActValue!=ReqValue) {
					Flag=true;
				}
				
			} else if(operator.equalsIgnoreCase("Greater than")){
				
				if (ActValue>ReqValue) {
					Flag=true;
				}
				
			} else if(operator.equalsIgnoreCase("Less than")){
			
				if (ActValue<ReqValue) {
					Flag=true;
				}
			}
			
			if(Flag!=true){
				resultmssg = "Exp : "+ActValue+" should "+operator+" act value : "+ReqValue+" : fail";
				exe_rst_status = 2;
			} else{
			    exe_rst_status = 1;
				resultmssg = "Exp : "+ActValue+" should "+operator+" act value : "+ReqValue+" : Pass";
			}
			
		} else {
			
			if(operator.equalsIgnoreCase("Equal to")){
				if (ActualValue.equalsIgnoreCase(Value)) {
					Flag=true;
				}
			}
			else if(operator.equalsIgnoreCase("Not Equal to")){
				if (!(ActualValue.equalsIgnoreCase(Value))) {
					Flag=true;
				}
			}
			
			if(Flag!=true){
				 resultmssg = "Exp "+Value+" should "+operator+" act value "+ActualValue+" : fail";
				exe_rst_status = 2;
			} else{
			    exe_rst_status = 1;
			    resultmssg = "Exp "+Value+" should "+operator+" act value "+ActualValue+" : Pass";
			}
		}
			
	}
	
	public static void Validate() {
		resultmssg = null ;
		exe_rst_status = 2;
		String PageTitle= driver.getTitle().trim();
		if (PageTitle.equalsIgnoreCase(Driver.ScreenName.trim())){
			exe_rst_status = 1;
			resultmssg = "Exp : "+PageTitle+" and actual value : "+Driver.ScreenName.trim()+" value matched!!";
		} else{
			exe_rst_status = 2;
			resultmssg = "Exp : "+PageTitle+" and actual value : "+Driver.ScreenName.trim()+" value not matched!!";
		}	
		
	}
	
	//This function just checks if the element passed as parameter is present or not
    public static boolean validateElementPresence(String objnm){
    	 boolean rs = false;
    	 resultmssg = null ;
		 exe_rst_status = 2;
		 //this function first checks if the element is present in xml file or not.
		 //if present in xml file it then checks what identification property it uses by
		 //checking the subtag in xml file that has a value. E.g class, id, xpath etc
		 //and finally finds the object on the application using FindelementBy. the property
		 //chosem
    	 List<WebElement> Object_name = GenericFunctions.Field_objs(objnm);
    	 //if element is found then return status is 1 and message is element present
    	 //otherwise return status is 2 and message is element not present
    	 if(Object_name.size()>0){
    		 exe_rst_status = 1;
 			resultmssg = "Element present";
 			rs = true;
    	 }else{
    		 exe_rst_status = 2;
  			resultmssg = "Element not present";
    	 }
    	 //return status as true or null
    	 return rs;
    }
	
    //This function accepts the String value PageObject as parameter and returns the value of the tagname 
    //and tag value in a hashmap
		public static  HashMap<String, String> XMLReading(String OR_Name) {
			
			//this get the value of the folder that stores all the xml files e.g BenefitAsia2_EE
			String XMLName = Driver.ORXMLName;	
			//gets the value of the actual xml file to be opened
		    String tempDir =Resourse_path.currPrjDirpath +"/ObjectRepositories/"+ Driver.ORNAME_XML +"/"+XMLName+".xml"; 
		    Logger.info("Xml obj_repo path "+tempDir);
		    //creates a new hashmap
			HashMap<String, String> map = new HashMap<String, String>();
			
			//checks is the xml file exists
			try {
				File File_obj = new File(tempDir);
				if(!File_obj.exists()){
					Logger.info(tempDir+" does not exist!!");
				}
				//gets the page object tag from the xml file
				DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
				Document doc = dBuilder.parse(File_obj);
				doc.getDocumentElement().normalize();
				NodeList nList = doc.getElementsByTagName(OR_Name);
				//validating tag in xml
				if(nList.getLength()<1){
					Logger.info("Defined excel object tag is not found in xml!!");
				}
	            //reading xml node
				//gets the value of the page object tag and stores in hashmap the tagname and tagvalue
				for (int temp = 0; temp < nList.getLength(); temp++) {
					Node nNode = nList.item(temp);
					if (nNode.getNodeType() == Node.ELEMENT_NODE) {
						Element eElement = (Element) nNode;
						NodeList ChildnodeList = eElement.getChildNodes();
						int lenghth = ChildnodeList.getLength();
						for (int j = 0; j < lenghth; j++) {
							Node Cnode = ChildnodeList.item(j);
							if (Cnode.getNodeType() == Node.ELEMENT_NODE) {
								String TagValue = Cnode.getTextContent().trim();
								String TagName = Cnode.getNodeName().trim();
								map.put(TagName.toLowerCase(), TagValue);
							}
						}
					}
				}
			} 
			
			//catch block to catch exception and log error
			catch (Exception e) {
	           e.printStackTrace();
	           Logger.error(e);
			}
			//return the map with tag name and tag value
			return map;
	    }

	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	//
	public static void windowInstancehandler(String OR_ObjectName){
		try{
			
			//create a new hashmap hp
			HashMap<String,String> hp = new HashMap<String,String>();
			//OR_Object name has the static value PageObject. The function XMLReading returns the tag name 
			// and tag value in a hashmap and stores in hashmap hp
			hp = XMLReading(OR_ObjectName);
			//if value is returned then get the value of the title tag 
			if(hp.size()>0){
				String data = hp.get("title");
				char [] chlength = data.toCharArray();
				//if value of title is not blank then
				if(chlength.length>0){
					
					handleMultipleWindows(data);
				}else{
					Logger.info("Title contains null value in xml!!");
					//function to switch window based on object
					handleMultipleWindowsObjectBased();
				}
			}else{
				Logger.info("Object is not defined in xml!!");
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	//this function creates an arraylist called localobjstorg and stores all the possible
	//identification types or tagnames in excel in the arraylist like id, xpath etc
	public static ArrayList<String> localObjStorage(){
		ArrayList<String> localobjstorg = new ArrayList<String>();
		// Mapped element
 		localobjstorg.add("id");
 		localobjstorg.add("class");
 		localobjstorg.add("xpath");
 		localobjstorg.add("name");
 		localobjstorg.add("linktext");
 		localobjstorg.add("tagname");
 		localobjstorg.add("css");
 		localobjstorg.add("partiallinktext");
 		
		return localobjstorg;
	}
	
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 * 
	 * 
	 */
	//creates a hashmap with all the properties that can be used for object identification
	public static HashMap<String, Integer> switchPrpStorage() {
		HashMap<String, Integer> swtich_prp_storage = new HashMap<String, Integer>();
		// Element mapped for executing operation
		swtich_prp_storage.put("id", 1);
		swtich_prp_storage.put("xpath", 2);
		swtich_prp_storage.put("class", 3);
		swtich_prp_storage.put("name", 4);
		swtich_prp_storage.put("linktext", 5);
		swtich_prp_storage.put("tagname", 6);
		swtich_prp_storage.put("css", 7);
		swtich_prp_storage.put("partiallinktext",8);
		
		return swtich_prp_storage;
	}
    
/** @author - ashish-choudhary
    @Function_Name -  Field_obj()
    @Description - It fetch all web object property value from xml, based on object and return the intended 
    webelement for executing the operation.
    @return WebElemet object
    @Created_Date - 18 Dec 2014
    @Modified_Date - 
 */
 	public static WebElement Field_obj(String OR_ObjectName) {
 		driverCheck();
 		WebElement WebEle_obj = null;
 		HashMap<String, String> xml_property_storage = XMLReading(OR_ObjectName);

 		// Storing property name to arraylist
 		ArrayList<String> prpstorage = new ArrayList<String>();

 		for (Map.Entry<String, String> m : xml_property_storage.entrySet()) {
 			prpstorage.add(m.getKey());
 		}

 		ArrayList<String> localobjstorg = localObjStorage();

 		String obj = null;
 		boolean exitprcess = false;
 		boolean mapped = false;
 		for (int i = 0; i < prpstorage.size(); i++) {
 			obj = prpstorage.get(i);
 			// Logger.info("first "+obj);
 			for (int j = 0; j < localobjstorg.size(); j++) {
 				String obj2 = localobjstorg.get(j);
 				if (obj.equalsIgnoreCase(obj2)) {
 					mapped = true;
 					Logger.info(obj+" object successfully mapped!!");
 					String objvalue = xml_property_storage.get(obj.toLowerCase());
 					char[] ch = objvalue.toCharArray();
 					if (ch.length > 0) {
 						Logger.info(obj+" object contains property value !!");
 						exitprcess = true;
 						break;
 					}else{
 						Logger.info(obj+" object contains null property value, so moving to next object..");
 					}
 				}
 			}
 			if (exitprcess != false) {
 				break;
 			}
 		}

 		// object code validation
 		if (exitprcess != false) {

 			String prp_value = xml_property_storage.get(obj);
 			Logger.info("Property name - " + obj);
 			Logger.info("Property value - " + prp_value);

 			HashMap<String, Integer> swtich_prp_storage = switchPrpStorage();

 			// storing property name to hashmap
 			ArrayList<String> switchmapobject = new ArrayList<String>();
 			for (Map.Entry<String, Integer> m : swtich_prp_storage.entrySet()) {
 				switchmapobject.add(m.getKey());
 			}
 			// validating mapping with switch case
 			boolean switchprocess = false;
 			for (int i = 0; i < switchmapobject.size(); i++) {
 				String mpobject = switchmapobject.get(i);
 				if (obj.equalsIgnoreCase(mpobject)) {
 					switchprocess = true;
 					break;
 				}
 			}

 			// swtich case validation
 			if (switchprocess != false) {
 				int switchid = swtich_prp_storage.get(obj);
 //				Logger.info(obj + " element is returned!!");
 				Logger.info("Finding "+obj+" element on webpage..");
 				
 				HashMap<String, String> hp1;
				// Assigning element to webelement
 				switch (switchid) {
 				case 1:
 					WebEle_obj = driver.findElement(By.id(prp_value));
 					break;
 				case 2:
 					String TC_Name = Driver.TestCaseID;
					 
 					/*
 					
 					if(TC_Name.equalsIgnoreCase("TC1"))
				 
					{	
						 	WebDriverWait wait = new WebDriverWait(driver, 600);
	 						wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
	 						WebEle_obj = driver.findElement(By.xpath(prp_value));
	 						break;
					}
					 else
	 					{
	 					
	 						//WebDriverWait wait = new WebDriverWait(driver, 90);
	 						//wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
	 						WebEle_obj = driver.findElement(By.xpath(prp_value));
	 						break;
	 					}
	 					
	 					*/
 					String TC_Name1 = Driver.TestCaseID;
					 if(TC_Name1.equalsIgnoreCase("TC1"))
				 
					{	
						WebDriverWait wait = new WebDriverWait(driver, 600);
						wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
						WebEle_obj = driver.findElement(By.xpath(prp_value));
						//returnPresenceofElement(WebEle_obj);
						break;
					}
					else
					{
					
						try {
							if ("1".equals(GenericFunctions.winCount()))
								{
									WebDriverWait wait = new WebDriverWait(driver, 90);
									wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
									WebEle_obj = driver.findElement(By.xpath(prp_value));
									//returnPresenceofElement(fielobjs);
									break;
								}
							else
							{
								WebEle_obj = driver.findElement(By.xpath(prp_value));
								//returnPresenceofElement(fielobjs);
								break;
								
							}
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						
					}
 					
 				case 3:
 					WebEle_obj = driver.findElement(By.className(prp_value));
 					break;
 				case 4:
 					WebEle_obj = driver.findElement(By.name(prp_value));
 					break;
 				case 5:
 					WebEle_obj = driver.findElement(By.linkText(prp_value));
 					break;
 				case 6:
 					WebEle_obj = driver.findElement(By.tagName(prp_value));
 					break;	
 				case 7:
 					WebEle_obj = driver.findElement(By.cssSelector(prp_value));
 					break;
 				case 8:
 					WebEle_obj = driver.findElement(By.partialLinkText(prp_value));
 					break;	
 				default:
 					Logger.info("Event is not mapped for this operation");
 					break;
 				}
 				Logger.info("Element found!!");
 			} else {
 				Logger.info(obj + " is not mapped in property local storage case!!");
 			}

 		} else {
 			if(xml_property_storage.size()>0){
 				if (mapped != true) {
 					Logger.info("Xml object is not mapped in local object storage!!");
 				}else {
 					if(exitprcess!=true){
 						Logger.info("Xml object is mapped but it does not contains object property value in xml file!!");
 					}
 				}
 			}else{
 				Logger.info("size of the element return by xmlreading is "+prpstorage.size());
 			}
 		}

 		return WebEle_obj;
 	}
 	
	/**
	 * @author - ashish-choudhary
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
 	
 	//The first fieldname in the test case sheet is passed as parameter
 	public static List<WebElement> Field_objs(String objectnm) {
 		
 		//checks if driver is null
 		driverCheck();
 		//create an arraylist to store webelements
 		List<WebElement> fielobjs = null;
 		//calls the XMLReading function and passes the fieldname as parameter. This function returns the tagname
 		//and tag value in a hashmap
 		HashMap<String, String> xml_property_storage = XMLReading(objectnm);

 		// creates a new arraylist prpstorage
 		ArrayList<String> prpstorage = new ArrayList<String>();

// 		Logger.info("prpstorage size "+xml_property_storage.size());
//		Logger.info("prpstorage "+xml_property_storage.get(0));
 		//Loops through the hashmap storing tagname and tagvalue and stores object name in 
 		//prpstorage arraylist
 		for (Map.Entry<String, String> m : xml_property_storage.entrySet()) {
 			String vl = m.getKey();
 			prpstorage.add(vl);
// 			Logger.info("Key "+vl);
 		}

 		//This function creates an arraylist which stores all the tagnames like id, xpath, class etc and 
 		//the arraylist returned is stored in the arraylist localobjstorg
 		ArrayList<String> localobjstorg = localObjStorage();

 		String obj = null;
 		boolean exitprcess = false;
 		boolean mapped = false;
 		//Loops through the arraylist prpstorage that has all the tag names or objects
 		for (int i = 0; i < prpstorage.size(); i++) {
 			//get the value of the object from the prpstorage arraylist to obj variable
 			obj = prpstorage.get(i);
 			//Loop through the array list which contains all the subtags like class, id, xpath etc
 			for (int j = 0; j < localobjstorg.size(); j++) {
 				//get the value of the tagname from the arraylist in a variable
 				String obj2 = localobjstorg.get(j);
 				//if both values match then set flag and log entry
 				if (obj.equalsIgnoreCase(obj2)) {
 					mapped = true;
 					Logger.info(obj+" object successfully mapped!!");
 					//get value of the tag
 					String objvalue = xml_property_storage.get(obj.toLowerCase());
 					char[] ch = objvalue.toCharArray();
 					//if value is not blank then set exitprocess flag to true and come out of loop
 					if (ch.length > 0) {
 						Logger.info(obj+" object contains property value !!");
 						exitprcess = true;
 						break;
 						//if not matching then log entry and continue
 					}else{
						Logger.info(obj+" object contains null property value, so moving to next object..");
 					}
 				}
 			}
 			//Come out of the main element if exit process flag is set to true
 			if (exitprcess != false) {
 				break;
 			}
 		}

 		// if exit process is true then get the value of property and value
 		if (exitprcess != false) {

 			String prp_value = xml_property_storage.get(obj);
 			Logger.info("Property name - " + obj);
 			Logger.info("Property value - " + prp_value);

 			//creates a hashmap with all the object identification properties that can be used
 			HashMap<String, Integer> swtich_prp_storage = switchPrpStorage();

 			// storing property name to arraylist
 			ArrayList<String> switchmapobject = new ArrayList<String>();
 			//Loop through the hashmap 
 			for (Map.Entry<String, Integer> m : swtich_prp_storage.entrySet()) {
 				//stores the key i.e the property type in arraylist switchmapobject
 				switchmapobject.add(m.getKey());
 			}
 			// validating mapping with switch case
 			boolean switchprocess = false;
 			for (int i = 0; i < switchmapobject.size(); i++) {
 				String mpobject = switchmapobject.get(i);
 				if (obj.equalsIgnoreCase(mpobject)) {
 					switchprocess = true;
 					break;
 				}
 			}

 			// swtich case validation
 			if (switchprocess != false) {
 				int switchid = swtich_prp_storage.get(obj);
// 				Logger.info(obj + " element is returned!!");
 				Logger.info("Finding "+obj+" element on webpage..");
 				//switch case to find element using the identification property mentioned in xml file
 				// Assigning element to webelement
 				switch (switchid) {
 				case 1:
 					fielobjs = driver.findElements(By.id(prp_value));
 					//this just checks if element size is >0 and does logging
 					returnPresenceofElement(fielobjs);
 					break;
 				case 2:
 				
 					String TC_Name = Driver.TestCaseID;
 					 if(TC_Name.equalsIgnoreCase("TC1"))
 				 
 					{	
 						WebDriverWait wait = new WebDriverWait(driver, 600);
 						wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
 						fielobjs = driver.findElements(By.xpath(prp_value));
 						returnPresenceofElement(fielobjs);
 						break;
 					}
 					else
 					{
 					
 						try {
							if ("1".equals(GenericFunctions.winCount()))
								{
									WebDriverWait wait = new WebDriverWait(driver, 90);
									wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(prp_value)));
									fielobjs = driver.findElements(By.xpath(prp_value));
									returnPresenceofElement(fielobjs);
									break;
								}
							else
							{
								fielobjs = driver.findElements(By.xpath(prp_value));
								returnPresenceofElement(fielobjs);
								break;
								
							}
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
 						
 					}
 				case 3:
 					fielobjs = driver.findElements(By.className(prp_value));
 					returnPresenceofElement(fielobjs);
 					break;
 				case 4:
 					fielobjs = driver.findElements(By.name(prp_value));
 					returnPresenceofElement(fielobjs);
 					break;
 				case 5:
 					WebDriverWait wait = new WebDriverWait(driver, 600);
					wait.until(ExpectedConditions.presenceOfElementLocated(By.linkText(prp_value)));
 					fielobjs = driver.findElements(By.linkText(prp_value));
 					returnPresenceofElement(fielobjs);
 					break;
 				case 6:
 					fielobjs = driver.findElements(By.tagName(prp_value));
 					returnPresenceofElement(fielobjs);
 					break;	
 				case 7:
 					fielobjs = driver.findElements(By.cssSelector(prp_value));
 					returnPresenceofElement(fielobjs);
 				case 8:
 					fielobjs = driver.findElements(By.partialLinkText(prp_value));	
 					returnPresenceofElement(fielobjs);
 					break;	
 				default:
 					Logger.info("Event is not mapped for this operation");
 					break;
 				}
 			} else {
 				Logger.info(obj + " is not mapped in property local storage!!");
 			}
 		} else {
 			if(xml_property_storage.size()>0){
 				if (mapped != true) {
 					Logger.info("Xml object is not mapped in local object storage!!");
 				}else {
 					if(exitprcess!=true){
 						Logger.info("Xml object is mapped but it does not contains object property value in xml file!!");
 					}
 				}
 			}else{
 				Logger.info("size of the element return by xmlreading is "+prpstorage.size());
 			}
 		}
 		//return the web element
 		return fielobjs;
 	}
 	
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
 	public static boolean dynamicwaitforObject(String objectnm, int time) {
 		driverCheck();
 		boolean fielobjs = false;
 		HashMap<String, String> xml_property_storage = XMLReading(objectnm);

 		// Storing property name to arraylist
 		ArrayList<String> prpstorage = new ArrayList<String>();

 		for (Map.Entry<String, String> m : xml_property_storage.entrySet()) {
 			String vl = m.getKey();
 			prpstorage.add(vl);
 //			Logger.info("Key "+vl);
 		}

 		ArrayList<String> localobjstorg = localObjStorage();

 		String obj = null;
 		boolean exitprcess = false;
 		boolean mapped = false;
 		for (int i = 0; i < prpstorage.size(); i++) {
 			obj = prpstorage.get(i);
 			// Logger.info("first "+obj);
 			for (int j = 0; j < localobjstorg.size(); j++) {
 				String obj2 = localobjstorg.get(j);
 				if (obj.equalsIgnoreCase(obj2)) {
 					mapped = true;
// 					Logger.info(obj+" object successfully mapped!!");
 					String objvalue = xml_property_storage.get(obj.toLowerCase());
 					char[] ch = objvalue.toCharArray();
 					if (ch.length > 0) {
 //						Logger.info(obj+" object contains property value !!");
 						exitprcess = true;
 						break;
 					}else{
	//					Logger.info(obj+" object contains null property value, so moving to next object..");
 					}
 				}
 			}
 			if (exitprcess != false) {
 				break;
 			}
 		}

 		// object code validation
 		if (exitprcess != false) {
 			String prp_value = xml_property_storage.get(obj);
 			HashMap<String, Integer> swtich_prp_storage = switchPrpStorage();
 			// storing property name to hashmap
 			ArrayList<String> switchmapobject = new ArrayList<String>();
 			for (Map.Entry<String, Integer> m : swtich_prp_storage.entrySet()) {
 				switchmapobject.add(m.getKey());
 			}
 			// validating mapping with switch case
 			boolean switchprocess = false;
 			for (int i = 0; i < switchmapobject.size(); i++) {
 				String mpobject = switchmapobject.get(i);
 				if (obj.equalsIgnoreCase(mpobject)) {
 					switchprocess = true;
 					break;
 				}
 			}

 			// swtich case validation
 			if (switchprocess != false) {
 				int switchid = swtich_prp_storage.get(obj);
 				Logger.info("Dynamic wait time for object is "+time+" sec");
 				// Assigning element to webelement
 				switch (switchid) {
 				case 1:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.id(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					break;
 				case 2:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.xpath(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					break;
 				case 3:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.className(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					break;
 				case 4:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.name(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					
 					break;
 				case 5:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.linkText(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					
 					break;
 				case 6:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.tagName(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					
 					break;	
 				case 7:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.cssSelector(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					
 					break;
 					
 				case 8:
 					for(int i=1; i<=time; i++){
 						List<WebElement> ele = driver.findElements(By.cssSelector(prp_value));
 						if(ele.size()>0){
 							Logger.info("Element present");
 							fielobjs = true;
 							break;
 						}else{
 							staticwait(1);
 							Logger.info("waiting for element to present !!"+i+" sec");
 						}
 					}
 					
 					break;	
 				default:
 					Logger.info("Event is not mapped for this operation");
 					break;
 				}
 			} else {
 				Logger.info(obj + " is not mapped in property local storage!!");
 			}
 			
 			if(fielobjs!=true){
 	 			Logger.info("Element is not present");
 	 		}
 			
 		} else {
 			if(xml_property_storage.size()>0){
 				if (mapped != true) {
 					Logger.info("Xml object is not mapped in local object storage!!");
 				}else {
 					if(exitprcess!=true){
 						Logger.info("Xml object is mapped but it does not contains object property value in xml file!!");
 					}
 				}
 			}else{
 				Logger.info("size of the element return by xmlreading is "+prpstorage.size());
 			}
 		}
 		return fielobjs;
 	}
 	
 	public static void returnPresenceofElement(List<WebElement> ele){
 		if(ele.size()>0){
 			Logger.info("Element found!!");
 		}else {
 			Logger.info("Element not found!!");
 		}
 	}
 	
	public static By returnByObject(String OR_ObjectName){
		driverCheck();
		By object = null;
		
 		HashMap<String, String> xml_property_storage = XMLReading(OR_ObjectName);

 		// Storing property name to arraylist
 		ArrayList<String> prpstorage = new ArrayList<String>();

 		for (Map.Entry<String, String> m : xml_property_storage.entrySet()) {
 			prpstorage.add(m.getKey());
 		}

 		ArrayList<String> localobjstorg = localObjStorage();

 		String obj = null;
 		boolean exitprcess = false;
 		boolean mapped = false;
 		for (int i = 0; i < prpstorage.size(); i++) {
 			obj = prpstorage.get(i);
 			// Logger.info("first "+obj);
 			for (int j = 0; j < localobjstorg.size(); j++) {
 				String obj2 = localobjstorg.get(j);
 				if (obj.equalsIgnoreCase(obj2)) {
 					mapped = true;
 					Logger.info(obj+" object successfully mapped!!");
 					String objvalue = xml_property_storage.get(obj.toLowerCase());
 					char[] ch = objvalue.toCharArray();
 					if (ch.length > 0) {
 						Logger.info(obj+" object contains property value !!");
 						exitprcess = true;
 						break;
 					}else{
 						Logger.info(obj+" object contains null property value, so moving to next object..");
 					}
 				}
 			}
 			if (exitprcess != false) {
 				break;
 			}
 		}

 		// object code validation
 		if (exitprcess != false) {

 			String prp_value = xml_property_storage.get(obj);
 			Logger.info("Property name - " + obj);
 			Logger.info("Property value - " + prp_value);

 			HashMap<String, Integer> swtich_prp_storage = switchPrpStorage();

 			// storing property name to hashmap
 			ArrayList<String> switchmapobject = new ArrayList<String>();
 			for (Map.Entry<String, Integer> m : swtich_prp_storage.entrySet()) {
 				switchmapobject.add(m.getKey());
 			}
 			// validating mapping with switch case
 			boolean switchprocess = false;
 			for (int i = 0; i < switchmapobject.size(); i++) {
 				String mpobject = switchmapobject.get(i);
 				if (obj.equalsIgnoreCase(mpobject)) {
 					switchprocess = true;
 					break;
 				}
 			}

 			// swtich case validation
 			if (switchprocess != false) {
 				int switchid = swtich_prp_storage.get(obj);
 //				Logger.info(obj + " element is returned!!");
 				Logger.info("Finding "+obj+" element on webpage..");
 				// Assigning element to webelement
 				switch (switchid) {
 				case 1:
 					object = By.id(prp_value);
 					break;
 				case 2:
 					object = By.xpath(prp_value);
 					break;
 				case 3:
 					object = By.className(prp_value);
 					break;
 				case 4:
 					object = By.name(prp_value);
 					break;
 				case 5:
 					object = By.linkText(prp_value);
 					break;	
 				case 6:
 					object = By.tagName(prp_value);
 					break;	
 				case 7:
 					object = By.cssSelector(prp_value);
 					break;	
 				case 8:
 					object = By.partialLinkText(prp_value);
 					break;	
 				default:
 					Logger.info("Event is not mapped for this operation");
 					break;
 				}
 				Logger.info("Element found!!");
 			} else {
 				Logger.info(obj + " is not mapped in property local storage!!");
 			}

 		} else {
 			if(xml_property_storage.size()>0){
 				if (mapped != true) {
 					Logger.info("Xml object is not mapped in local object storage!!");
 				}else {
 					if(exitprcess!=true){
 						Logger.info("Xml object is mapped but it does not contains object property value in xml file!!");
 					}
 				}
 			}else{
 				Logger.info("size of the element return by xmlreading is "+prpstorage.size());
 			}
 		}
		
		return object;
	}
	
	/**
	 * @author - ashish-choudhary
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	public static void OpenURL(String URL, String browserName) {
		exe_rst_status = 2;
		boolean openb = false;
		try {
			if (driver == null) {
				if ((browserName.equalsIgnoreCase("ff"))) {
					//System.setProperty("webdriver.firefox.bin", "C:\\Users\\Gaurav-Chhabra\\firefox.exe");
//					System.setProperty("webdriver.firefox.bin", "C:\\Users\\ashish-choudhary\\Downloads\\firefox.exe");
					Logger.info("please wait.. browser getting open !!");
					driver = new FirefoxDriver(FirefoxDriverProfile());
					//driver.get(URL);
					//Thread.sleep(1000 * 7);
					//driver.manage().window().maximize();
					//Logger.info("please wait.. browser getting open !!");
					//driver = new FirefoxDriver();
					openb = true;
				} else if (browserName.equalsIgnoreCase("htmlunit")){
					
					HtmlUnitDriver driver = new HtmlUnitDriver(true);
					
					driver.get("http://ausyd16as13v/BA2Web_WIP/ba2admin/Home/tabid/121/Default.aspx");
					//System.out.println(URL);
					((HtmlUnitDriver) driver).setJavascriptEnabled(true);
					//((HtmlUnitDriver) driver).setJavascriptEnabled(true);
					openb = true;
				} else if (browserName.equalsIgnoreCase("ie")) {
					String iedriverserver = Resourse_path.browserDriverpath;
					//System.setProperty("webdriver.ie.driver",iedriverserver);
					System.setProperty(InternetExplorerDriverService.IE_DRIVER_EXE_PROPERTY,iedriverserver);
					DesiredCapabilities capab = DesiredCapabilities.internetExplorer();
					capab.setCapability(CapabilityType.ForSeleniumServer.ENSURING_CLEAN_SESSION, true); 
			        capab.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
			        capab.setCapability("ignoreZoomSetting", true);
			        capab.setCapability("nativeEvents",false);
			        capab.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
					driver = new InternetExplorerDriver(capab);
					driver.manage().deleteAllCookies();
					openb = true;
				} else if (browserName.equalsIgnoreCase("chrome")) {
					String chromedriverserver = Resourse_path.chrome_driver_path;
					System.setProperty("webdriver.chrome.driver",chromedriverserver);
					driver = new ChromeDriver();
					openb = true;
				} else if (browserName.equalsIgnoreCase("auth_ff")) {
					driver = new FirefoxDriver();
					driver.manage().window().maximize();
					String ExePath = Resourse_path.currPrjDirpath+"\\BrowserDrivers\\Auth.exe";
					String Username = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,1, 10);
					String Password = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,1, 12);
					Logger.info("Username" + Username);
					Logger.info("Password" + Password);
					Process pb = new ProcessBuilder(ExePath, "Firefox", "30",Username, Password).start();
					driver.get(URL);
					Thread.sleep(1000 * 30);
					pb.destroy();
					Logger.info("Browser is opened with " + URL);
				} else {
					Logger.info("Failed: Invalid Browser: "+ browserName);
				}
			} else {
				Logger.info("Navigating to another url..");
				driver.manage().window().maximize();
				driver.navigate().to(URL);
			}
			// open browser
			if (openb == true) {
				Logger.info("Opening Application..");
				driver.manage().window().maximize();
				driver.get(URL);
				driver.manage().timeouts().implicitlyWait(40,TimeUnit.SECONDS);
				Logger.info("Browser is opened with " + URL);
			}
			exe_rst_status = 1;
		} catch (Exception e) {
			exe_rst_status = 2;
			e.printStackTrace();
			Logger.error(e);
			Driver.exceptionreport("Object should be found","Object not found!!",e.getMessage());
		}
	}
    
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	public static void executeRobotevent(String event_nm) {
		try {
			exe_rst_status = 2;
			Robot robot = new Robot();
			HashMap<String, Integer> map = new HashMap<String, Integer>();

			map.put("robot_tab", 1);
			map.put("robot_enter", 2);
			map.put("robot_down", 3);
			map.put("robot_right", 4);
			map.put("robot_typecharacter", 5);

			int dataa = map.get(event_nm.toLowerCase());
			String msg = "operation perfomred!!";

			switch (dataa) {
			case 1:
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				Logger.info(event_nm + " " + msg);
				exe_rst_status = 1;
				break;
			case 2:
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				Logger.info(event_nm + " " + msg);
				exe_rst_status = 1;
				break;
			case 3:
				robot.keyPress(KeyEvent.VK_DOWN);
				robot.keyRelease(KeyEvent.VK_DOWN);
				Logger.info(event_nm + " " + msg);
				exe_rst_status = 1;
				break;
			case 4:
				robot.keyPress(KeyEvent.VK_RIGHT);
				robot.keyRelease(KeyEvent.VK_RIGHT);
				Logger.info(event_nm + " " + msg);
				exe_rst_status = 1;
				break;
			default:
				Logger.info("Event is not mapped for this operation");
				exe_rst_status = 2;
				break;
			}
		} catch (Exception e) {
			exe_rst_status = 1;
			e.printStackTrace();
			Logger.error(e);
		}

	}
	
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	
	//this function accepts the value of field value as parameter and is used for clicking on
	//links or autoit or robot operations
	public static void OtherClickOperation(XSSFCell Cell_obj6){
		if(Cell_obj6!=null){
    		String data1 = Cell_obj6.toString();
			Logger.info("col 6 data : "+data1);
			char [] ch = data1.toCharArray();
			if(ch.length > 0){
				String act_name = Cell_obj6.getStringCellValue();
				if(act_name.contains("link")){
					Logger.info("under link operation");
					GenericFunctions.handlelinkOperation("Click",act_name);
				}else if(act_name.contains("autoit_")){
					Logger.info("under autoIt operation");
					GenericFunctions.executeAutoItOperation(act_name);
				}else if(act_name.contains("robot_")){
					Logger.info("under robot operation");
					GenericFunctions.executeRobotevent(act_name);
				}
			}
    	}
	}
	
	
	/**
	 * @author - ashish-choudhary
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	public static void executeActionOperation(String element, String subelement){
		exe_rst_status =2;
		Actions builder = new Actions(driver);
		WebElement Object_name = GenericFunctions.Field_obj(element);
		WebElement subElement = GenericFunctions.Field_obj(subelement);
		// operation execution
        builder.moveToElement(Object_name).perform();
        builder.moveToElement(subElement);
        builder.click();
        builder.perform();
		
		exe_rst_status = 1;
		
	}
	
	
	/**
	 * @author - brijesh-yadav
	 * @Function_Name -
	 * @Description -
	 * @return
	 * @Created_Date -
	 * @Modified_Date -
	 */
	public static void executeActionBuilderOperation(String event_nm, String element){
		exe_rst_status = 2;
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		Actions builder = new Actions(driver);
		// Action mapping
		map.put("act_single", 1);
		map.put("act_double", 2);
		map.put("act_xy0", 3);
		map.put("act_hover", 4);
		map.put("act_mouseclick", 5);
		map.put("act_cordinatesclick", 6);
		map.put("act_cordinates_mouseclick", 7);
		map.put("act_rightclick", 8);
		map.put("act_mouseclick_sunmenu", 9);
		
		WebElement Object_name = GenericFunctions.Field_obj(element);
		Point coordinates = Object_name.getLocation();
		int x = coordinates.getX();
		int y = coordinates.getY();
		Robot robot;
		// switching data operation
		int data = map.get(event_nm.toLowerCase());
		String msg = "click "+event_nm + " " +"operation perfomred!!";
		// operation execution
		switch (data) {
		case 1:
			Logger.info(msg);
			builder.moveToElement(Object_name).click().build().perform();
			exe_rst_status = 1;
			break;
		case 2:
			Logger.info(msg);
			builder.moveToElement(Object_name).doubleClick().build().perform();
			exe_rst_status = 1;
			break;
		case 3:
			Logger.info(msg);
			builder.moveToElement(Object_name,0, 0).click().build().perform();
			exe_rst_status = 1;
			break;
		case 4:
			Logger.info(msg);
			builder.moveToElement(Object_name).build().perform();
			exe_rst_status = 1;
			break;
		case 5:
			Logger.info(msg);
			try {
				robot = new Robot();
				robot.mouseMove(coordinates.getX()+230,coordinates.getY()+330); 
				robot.mousePress(InputEvent.BUTTON1_MASK );
				robot.mouseRelease( InputEvent.BUTTON1_MASK);
			} catch (AWTException e) {
				e.printStackTrace();
				Logger.error(e);
			}
			exe_rst_status = 1;
			break;	
		case 6:
			Logger.info(msg);
			builder.moveToElement(Object_name,x,y).click().build().perform();
			exe_rst_status = 1;	
			break;
		case 7:
			Logger.info(msg);
			try {
				robot = new Robot();
				String xcord = fn_Data("xcord");
				String ycord = fn_Data("ycord");
				if(xcord !=null && ycord!=null){
					int xcords = Integer.parseInt(xcord);
					int ycords = Integer.parseInt(ycord);
					robot.mouseMove(coordinates.getX()+xcords,coordinates.getY()+ycords); 
					robot.mousePress(InputEvent.BUTTON1_MASK );
					robot.mouseRelease( InputEvent.BUTTON1_MASK);
				}
				else{
					Logger.info("x and y cordinates are null!!");
				}
			} catch (AWTException e) {
				e.printStackTrace();
				Logger.error(e);
			}
			exe_rst_status = 1;	
			break;
		case 8:
			Logger.info(msg);
			builder.contextClick(Object_name).build().perform();
			exe_rst_status = 1;	
			break;
		default:
			Logger.info("Event is not mapped for this operation");
			break;
		}
	}
	
	// Execute action builder class operation
	//this method is executed if we are clicking on one object only but there is some value present in Fieldvalu
	//column also like javascript click or alert_accept etc it accepts the object name to be clicked and the
	//fieldvalue as parameter
	public static void executeclickOps(String objname,String event_nm) {
		
		exe_rst_status = 2;
		//if fieldvalue contains act_
		if(event_nm.contains("act_")){
			//this function performs the click operation using the actions class
			executeActionBuilderOperation(event_nm, objname);
		}
		//this is executed if we dont have to use actions class
		else
		{
			
			//create an hashmap object called map to map the actions like javascript click, alert_accept, 
			//autoit etc 
			HashMap<String, Integer> map = new HashMap<String, Integer>();
			// Action mapping
			map.put("javascript_click", 1);
			map.put("alert_accept", 2);
			map.put("alert_dismiss", 3);
			map.put("conditional_click", 4);
			map.put("autoit_ff_openfile", 5);
			map.put("autoit_ie_openfile", 6);
			map.put("download_file", 7);
			map.put("autoit_ie_savefile_2", 8);
			map.put("autoit", 9);
			map.put("new_autoit", 10);
			map.put("cordinates_mousedouble_click", 11);
			
			// switching data operation
			String oprnm = event_nm.toLowerCase().trim();
			Logger.info("oprnm "+oprnm);
			
			int data = 0;
			if (map.containsKey(oprnm)){
				 data = map.get(oprnm);
			}
			String filename = fn_Data("filename");
			String msg = "click "+ event_nm + " " +"operation perfomred!!";
			// operation execution for diffetent types of clicks
			switch (data) {
			case 0:
				Logger.info(oprnm+" key does not match!!");
				exe_rst_status = 2;
				break;
			case 1:
				Logger.info(msg);
				JavascriptExecutor js = (JavascriptExecutor) driver;
				WebElement upload = GenericFunctions.Field_obj(Driver.OR_ObjectName);
				js.executeScript("arguments[0].click()", upload);
				System.out.println("autoit part");
				exe_rst_status = 1;
				String action = fn_Data("action");
				if(action!=null){
					char [] lngh = action.toCharArray();
					Logger.info("action "+ action);
					if(lngh.length>0){
						staticwait(12);
						executeAutoItOperation(action);
						staticwait(9);
					}
				}
				
				break;
			case 2:
				WebElement Object_name = GenericFunctions.Field_obj(objname);
				Logger.info(msg);
				Object_name.click();
			    if(isAlertPresent()==true){
			    	staticwait(1);
			    	driver.switchTo().alert().accept();
			    	if(isAlertPresent()==true){
				    	staticwait(1);
				    	driver.switchTo().alert().accept();
				    }
			    }
				exe_rst_status = 1;
				break;
			case 3:
				WebElement Object_name2 = GenericFunctions.Field_obj(objname);
				Logger.info(msg);
				Object_name2.click();
			    if(isAlertPresent()==true){
			    	staticwait(1);
			    	driver.switchTo().alert().dismiss();
			    }
				exe_rst_status = 1;
				break;
			case 4:
				exe_rst_status = 0;
				Logger.info(msg);
				List<WebElement> ele = Field_objs(objname);
				if(ele.size()>0){
					Click(objname);
					Logger.info("Conditional element is clicked!!");
					exe_rst_status = 1;
				}
				break;
			case 5:
				Logger.info(msg);
				//delte file before open if it exist..
				deleteFile(filename);
				Click(objname);
				staticwait(5);
				executeAutoItOperation(event_nm.toLowerCase());
				staticwait(5);
				exe_rst_status = 1;
				break;
			case 6:
				Logger.info(msg);
				//delte file before open if it exist..
				deleteFile(filename);
				Click(objname);
				staticwait(12);
				executeAutoItOperation(event_nm.toLowerCase());
				staticwait(9);
				exe_rst_status = 1;
				break;
			case 7:
				Logger.info(msg);
				downloadFile();
				break;
			case 8:
				Logger.info(msg);
				//delte file before open if it exist..
				deleteFile(filename);
				Click(objname);
				staticwait(12);
				executeAutoItOperation(event_nm.toLowerCase());
				staticwait(9);
				exe_rst_status = 1;
			case 9:
				Logger.info(msg);
				//delte file before open if it exist..
				deleteFile(filename);
				Click(objname);
				staticwait(12);
				executeAutoItOperation(event_nm.toLowerCase());
				staticwait(9);
				exe_rst_status = 1;	
				break;
			case 10:
				Logger.info(msg);
				//delte file before open if it exist..
				deleteFile(filename);
				Click(objname);
				staticwait(10);
				executeAutoItOperation(event_nm.toLowerCase());
				staticwait(12);
				exe_rst_status = 1;	
				break;	
			case 11:
				Logger.info(msg);
				WebElement Obj_nm = GenericFunctions.Field_obj(objname);
				Point coordinates = Obj_nm.getLocation();
				try {
					Robot robot = new Robot();
					String xcord = fn_Data("xcord");
					String ycord = fn_Data("ycord");
					if(xcord !=null && ycord!=null){
						int xcords = Integer.parseInt(xcord);
						int ycords = Integer.parseInt(ycord);
						robot.mouseMove(coordinates.getX()+xcords,coordinates.getY()+ycords); 
						robot.mousePress(InputEvent.BUTTON1_MASK );
						robot.mouseRelease( InputEvent.BUTTON1_MASK);
						robot.mousePress(InputEvent.BUTTON1_MASK );
						robot.mouseRelease( InputEvent.BUTTON1_MASK);
						exe_rst_status = 1;	
					}
					else{
						Logger.info("x and y cordinates are null!!");
					}
				} catch (AWTException e) {
					e.printStackTrace();
				}
				break;	
				
			default:
				Logger.info("Event is not mapped for this operation");
				break;
			}
		}
	}
	
	// get any value from excel file based on row and column
	public static String getValuefromExcel(String filepath, String sheetname,int row, int column) {
		String colValue = "none";
		try {
			FileInputStream FIS = new FileInputStream(filepath);
			XSSFWorkbook Workbook_obj = new XSSFWorkbook(FIS);
			XSSFSheet sheet_obj = Workbook_obj.getSheet(sheetname);
			XSSFRow row_obj = sheet_obj.getRow(row);
			XSSFCell cell_obj = row_obj.getCell(column);

			if (cell_obj != null) {
				colValue = cell_obj.toString();
			}
			Workbook_obj.close();
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
			colValue = "none";
		}
		return colValue;
	}
	
 //This function checks if there is a 
	public static boolean isRowEmptyInExcel(Row row) {
		try {
			for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
				Cell cell = row.getCell(c);
				if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
					return false;
				}
			}
		} catch (Exception e) {
			
		}
		return true;
	}
	
	
	//thjis functions tries to find the object specified as parameter
	public static void waitFortheElement(int waitTime, String element) {
		WebElement Object_name = GenericFunctions.Field_obj(element);
		WebDriverWait wait = new WebDriverWait(driver, waitTime); // The int here is the maximum time in seconds the element can wait.
		wait.until(ExpectedConditions.visibilityOf(Object_name));
	}
	
	
	// Executing wait operation, accepts the fieldname1 and fieldvalue 1 as parameter
	public static boolean waitOperation(XSSFCell obj_5, XSSFCell col_6) {
		boolean process_analyzer = false;
		exe_rst_status = 2;
		resultmssg = null ;
		// if fieldvalue is specified as a paramter
		if(col_6!=null){
			
			//if fieldvalue has some value
			boolean celstatus = validateCell(col_6);
			if(celstatus!=false)
			{
				XSSFCell Cell_objname = col_6;
				String keyword_nm = Cell_objname.getStringCellValue();
				//get the value from fieldvalue column
				// static and dynamic wait operation
				//if fieldvalue has the value dynamic wait
				if (keyword_nm.equalsIgnoreCase("dynamicWait")) 
				{
					//get the value of fieldname
					XSSFCell Cell_objname2 = obj_5;
					if (Cell_objname2 != null)
					{  
						//
						String ObjectName = Cell_objname2.getStringCellValue();
						char[] celllength = ObjectName.toCharArray();
						if (celllength.length > 0)
						{
							//if fieldname has some value
							//run this loop 420 times
							for (int i = 1; i < 420; i++) {
								Logger.info("ObjectName "+ObjectName);
								//get the webelement object for the object specified in the fieldname column
								List<WebElement> Object_name = GenericFunctions.Field_objs(ObjectName);
								Logger.info("Size of object "+ Object_name.size());
								//if webelement object is present then set status as true and do logging and 
								//set a flag and come out of the loop
								if (Object_name.size() > 0) {
									process_analyzer = true;
									exe_rst_status = 1;
									resultmssg = "Element found !!";
									Logger.info(resultmssg);
									break;
								}
								//if object name is not found
								else
								{
									try {
										//wait for one second
										Thread.sleep(1000);
									}
									//catch block for waiting 1 sec
									catch (InterruptedException e)
									{
										e.printStackTrace();
										Logger.error(e);
									}
									
									if (i % 10 == 0) {
										Logger.info("System is waiting for the element..");
									}
								}
							}
							//if after running the loop for 420 seconds i.e 7 minutes the object is not found then
							//set status as fail and do logging
							if(process_analyzer!=true){
								resultmssg = "Element is not found !!";
								exe_rst_status = 2;
								Logger.info(resultmssg);
							}
						}
						//if object name is not specified
						else {
							resultmssg = keyword_nm+ " Cell contains null value!!";
							exe_rst_status = 2;
							Logger.info(resultmssg);
						}
					} 
					//if object name is not specified
					else
					{
						resultmssg = keyword_nm + " Cell contains null value!!";
						exe_rst_status = 2;
						Logger.info(resultmssg);
					}
				}
			}
			
			else{
				process_analyzer = true;
				int time = Integer.parseInt(obj_5.toString());
				Logger.info("Waited for the element for "+time+" sec");
				try{
					Logger.info("waiting for the element for "+time+" sec");
					for(int i=time; i>=1; i--){
						Thread.sleep(1000);
						if(i%5==0){
							Logger.info("Time remaining "+i+" sec");
						}
					}
				}catch(Exception e){
					e.printStackTrace();
					Logger.error(e);
				}
				resultmssg = "Waited for the element for "+time+" sec";
				exe_rst_status = 1;
			}
		}else {
			process_analyzer = true;
			int time = Integer.parseInt(obj_5.toString());
			Logger.info("Waited for the element for "+time+" sec");
			try{
				Logger.info("waiting for the element for "+time+" sec");
				for(int i=time; i>=1; i--){
					Thread.sleep(1000);
					if(i%5==0){
						Logger.info("Time remaining "+i+" sec");
					}
				}
			}catch(Exception e){
				e.printStackTrace();
				Logger.error(e);
			}
			resultmssg = "Waited for the element for "+time+" sec";
			exe_rst_status = 1;
		}
		return process_analyzer;
	}
	
	public static String returnTotalTime(String start_time, String end_time) {
		String time = null;
		try {
			// String time1 = "16:40:30";
			// String time2 = "16:58:20";
			SimpleDateFormat format = new SimpleDateFormat("HH:mm:ss");
			Date date1 = format.parse(start_time);
			Date date2 = format.parse(end_time);
			long difference = date2.getTime() - date1.getTime();
			long milliseconds = difference;
            //divide the time
			int seconds = (int) (milliseconds / 1000) % 60;
			int minutes = (int) ((milliseconds / (1000 * 60)) % 60);
			int hours = (int) ((milliseconds / (1000 * 60 * 60)) % 24);
			
//			Logger.info(hours + ":" + minutes + ":" + seconds);
			time = hours + ":" + minutes + ":" + seconds;
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
		return time;
	}
	
	//This function accepts a string splits it using comma as the seperator and returns the seperated elements 
	//as a arraylist
	public static ArrayList<String> returnArraylistStringCommaSeprated(String data){
		ArrayList<String> arr = new ArrayList<String>();
		try{
			List<String> list = new ArrayList<String>(Arrays.asList(data.split(",")));
			//adding data to arraylist
			for(int i=0; i<list.size(); i++){
				arr.add(list.get(i).trim());
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
		return arr;
	}
	

	//this function is for dynamic operations like multiple clicks in one operation it accepts
	//parameter operation name i.e values like click and suboperationame e.g textbased
	public static void handleDynamicOperation(String operationname, String subOperationName){
		try{
			
			HashMap<String, Integer> map = new HashMap<String, Integer>();
			// Action mapping
			map.put("click", 1);
			map.put("validate", 2);
			map.put("input", 3);
			
			// switching data operation
			int data = map.get(operationname.toLowerCase());
			String msg = "operation perfomred!!";
			
			switch(data){
			case 1:
				Logger.info(operationname + " " + msg);
				//this function is called for value 1 i.e click and passes the suboperation name as parameter
				//like textbased
				clickDynacmicOperation(subOperationName);
				break;
			case 2:
				Logger.info(operationname + " " + msg);
				validateDynacmicOperation(subOperationName);
				break;
			case 3:
				Logger.info(operationname + " " + msg);
				break;
			case 4:
				Logger.info(operationname + " " + msg);
				break;
			default:
				Logger.info(operationname+" case is not mapped for this operation");
				break;
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	public static void validateDynacmicOperation(String valdateOperationName){
		
		ArrayList<String> arr = new ArrayList<String>();
		arr = GenericFunctions.returnArraylistStringCommaSeprated(Driver.OR_ObjectName);
		
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		// Action mapping
		map.put("text", 1);
		map.put("text_xpath", 2);
		map.put("text_xpath_add", 3);
		map.put("element", 4);
		
		// switching data operation
		int data = map.get(valdateOperationName.toLowerCase());
		String msg = "operation perfomred!!";
		String finalmsg = valdateOperationName + " " + msg ;
		
		switch(data){
		case 1:
			Logger.info(finalmsg);
			validateTextDynacmicOperationCases(arr);
			break;
		case 2:
			Logger.info(finalmsg);
			validateTextKeyValue(arr);
			break;
		case 3:
			Logger.info(finalmsg);
			break;
		case 4:
			Logger.info(finalmsg);
			break;
		default:
			Logger.info(valdateOperationName+" case is not mapped for this operation");
			break;	
		}
	}
	
	public static void validateTextDynacmicOperationCases(ArrayList<String> list){
		
		int operation_switch = list.size();
		String msg = "operation perfomred!!";
		String finalmsg = operation_switch + " " + msg ;
		
		int rownm = Driver.executing_row;
		String fieldValue_1 = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
		
		switch(operation_switch){
		case 2:
			Logger.info(finalmsg);
			validateTextDynamicOperation2(list,fieldValue_1);
			break;
		case 3:
			Logger.info(finalmsg);
			validateTextDynamicOperation3(list,fieldValue_1);
			break;
		case 4:
			Logger.info(finalmsg);
			break;
		case 5:
			Logger.info(finalmsg);
			break;
		default:
			Logger.info(operation_switch+" case is not mapped for this operation");
			break;	
		}
	}

	public static void validateTextDynamicOperation2(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				actData = obj2.getText();
				Logger.info("actData "+actData);
				if(actData.equalsIgnoreCase(expData)){
					rst = true;
					Logger.info("Actual data : "+actData+" and "+" expected data"+ expData+" are equal" );
					exe_rst_status = 1;
					break;
				}
			}
		}
		
		if(rst!=true){
			exe_rst_status = 2;
			Logger.info("Actual data : "+actData+" and "+" expected data"+ expData+" are equal" );
		}
	}
	
	public static void validateTextDynamicOperation3(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		String element_3 = list.get(2);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		By by3 = returnByObject(element_3);
		
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				List<WebElement> listobj2 = obj2.findElements(by3);
				if(listobj2.size()>0){
					for(int loop2=0; loop2<listobj2.size(); loop2++){
						WebElement obj3 = listobj2.get(loop2);
						actData = obj3.getText();
						Logger.info("actData "+actData);
						if(actData.equalsIgnoreCase(expData)){
							rst = true;
							Logger.info("Actual data : "+actData+" and "+" expected data"+ expData+" are equal" );
							exe_rst_status = 1;
							break;
						}
					}
				}
				if(rst!=false){
					break;
				}
			}
		}
		
		if(rst!=true){
			exe_rst_status = 2;
			Logger.info("Actual data : "+actData+" and "+" expected data"+ expData+" are equal" );
		}
		
	}
	
	public static void validateTextKeyValue(ArrayList<String> list){
		exe_rst_status = 2;
		int rownm = Driver.executing_row;
		String keyName = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 7);
		String value  = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
		
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		
		HashMap<String, String> value1 = XMLReading(element_1);
		String xpath1 = value1.get("xpath");
        
        HashMap<String, String> value2 = XMLReading(element_2);
        String xpath2 = value2.get("xpath");
        boolean keyStatus =false;
        boolean valueStatus =false;
        for(int i=1; i<60; i++){
        	int j = i ;
        	String finalelement = xpath1+j+xpath2;
        	List<WebElement> listele = driver.findElements(By.xpath(finalelement));
        	if(listele.size()>0){
        		String eledata = driver.findElement(By.xpath(finalelement)).getText();
        		Logger.info("keyName "+eledata);
        		if(eledata.equalsIgnoreCase(keyName)){
        			keyStatus = true;
        			int add_1 = j+1;
        			String finalelement2 = xpath1+add_1+xpath2;
                	List<WebElement> listele2 = driver.findElements(By.xpath(finalelement2));
                	if(listele2.size()>0){
                		String eledata2 = driver.findElement(By.xpath(finalelement2)).getText();
                		Logger.info("value "+eledata2);
                		if(eledata2.equalsIgnoreCase(value)){
                			valueStatus = true;
                			Logger.info("keyName "+keyName +" and keyName value "+value+" is matched !!");
                			exe_rst_status = 1;
                			break;
                		}
                	}
                	break;
        		}
        	}
        	if (i % 10 == 0) {
				Logger.info("System is searching for the element..");
			}
        }
        
        if(keyStatus!=true){
        	exe_rst_status = 2;
        	Logger.info("keyName "+keyName +" is not found!!");
        }else{
        	if(valueStatus!=true){
        		exe_rst_status = 2;
        		Logger.info("keyName value "+value +" is not matched!!");
        	}
        }
	}
	//for performing dynamic click operation with value of suboperationname like textbased as parameter
	public static void clickDynacmicOperation(String valdateOperationName){
		
		//creates a new array list
		ArrayList<String> arr = new ArrayList<String>();
		arr = GenericFunctions.returnArraylistStringCommaSeprated(Driver.OR_ObjectName);
		
		HashMap<String, Integer> map = new HashMap<String, Integer>();
		// Action mapping
		map.put("textbased", 1);
		map.put("hover_textbased", 2);
		
		// switching data operation
		int data = map.get(valdateOperationName.toLowerCase());
		String msg = "operation perfomred!!";
		String finalmsg = valdateOperationName + " " + msg ;
		
		switch(data){
		case 1:
			Logger.info(finalmsg);
			//This is for executing textbased dynamic click
			clickTextbasedDynacmicOperationCases(arr);
			break;
		case 2:
			Logger.info(finalmsg);
			hoverTextbasedDynacmicOperationCases(arr);
			break;
		case 3:
			Logger.info(finalmsg);
			hoverTextbasedDynacmicOperationCases(arr);
			break;	
		default:
			Logger.info(valdateOperationName+" case is not mapped for this operation");
			break;	
		}
	}
	
	public static void clickTextbasedDynacmicOperationCases(ArrayList<String> list){
		
		int operation_switch = list.size();
		String msg = "operation perfomred!!";
		String finalmsg = operation_switch + " " + msg ;
		int rownm = Driver.executing_row;
		
		String fieldValue_1 = getValuefromExcel(Driver.TestData_Sheetpath, "Test Case",rownm, 8);
		
		switch(operation_switch){
		case 2:
			Logger.info(finalmsg);
			ClickTextbasedDynamicOperation2(list,fieldValue_1);
			break;
		case 3:
			Logger.info(finalmsg);
			ClickTextbasedDynamicOperation3(list,fieldValue_1);
			break;
		case 4:
			Logger.info(finalmsg);
			ClickTextbasedDynamicOperation4(list,fieldValue_1);
			break;
		case 5:
			Logger.info(finalmsg);
			break;
		default:
			Logger.info(operation_switch+" case is not mapped for this operation");
			break;	
		}
	}

	public static void hoverTextbasedDynacmicOperationCases(ArrayList<String> list){
		
		int operation_switch = list.size();
		String msg = "operation perfomred!!";
		String finalmsg = operation_switch + " " + msg ;
		int rownm = Driver.executing_row;
		
		String fieldValue_1 = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
		
		switch(operation_switch){
		case 2:
			Logger.info(finalmsg);
			ClickTextbasedDynamicOperation2(list,fieldValue_1);
			break;
		case 3:
			Logger.info(finalmsg);
			ClickTextbasedDynamicOperation3(list,fieldValue_1);
			break;
		case 4:
			Logger.info(finalmsg);
			hoverTextbasedDynamicOperation4(list,fieldValue_1);
			break;
		case 5:
			Logger.info(finalmsg);
			break;
		default:
			Logger.info(operation_switch+" case is not mapped for this operation");
			break;	
		}
	}

	
	public static void ClickTextbasedDynamicOperation2(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				actData = obj2.getText();
				Logger.info("actData "+actData);
				if(actData.equalsIgnoreCase(expData)){
					rst = true;
					obj2.click();
					exe_rst_status = 1;
					break;
				}
				
			}
		}
		
		if(rst!=true){
			exe_rst_status = 2;
			Logger.info("expData "+expData+"is not found !!" );
		}
	}
	
	public static void ClickTextbasedDynamicOperation3(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		String element_3 = list.get(2);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		By by3 = returnByObject(element_3);
		
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				List<WebElement> listobj2 = obj2.findElements(by3);
				if(listobj2.size()>0){
					for(int loop2=0; loop2<listobj2.size(); loop2++){
						WebElement obj3 = listobj2.get(loop2);
						actData = obj3.getText();
						Logger.info("actData "+actData);
						if(actData.equalsIgnoreCase(expData)){
							rst = true;
							obj3.click();
							exe_rst_status = 1;
							break;
						}
					}
				}
				if(rst!=false){
					break;
				}
			}
		}
		
		if(rst!=true){
			Logger.info("expData "+expData+"is not found !!");
			exe_rst_status = 2;
		}
	}
	
	public static void ClickTextbasedDynamicOperation4(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		String element_3 = list.get(2);
		String element_4 = list.get(3);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		By by3 = returnByObject(element_3);
		By by4 = returnByObject(element_4);
		
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				List<WebElement> listobj2 = obj2.findElements(by3);
				if(listobj2.size()>0){
					for(int loop2=0; loop2<listobj2.size(); loop2++){
						WebElement obj3 = listobj2.get(loop2);
						List<WebElement> listobj3 = obj3.findElements(by4);
						if(listobj3.size() > 0){
							for(int loop3=0; loop3<listobj3.size(); loop3++){
								WebElement obj4 = listobj3.get(loop3);
								actData = obj4.getText();
								Logger.info("actData "+actData);
								if(actData.equalsIgnoreCase(expData)){
									rst = true;
									obj4.click();
									exe_rst_status = 1;
									break;
								}
							}
						}
						if(rst!=false){
							break;
						}
					}
				}
				if(rst!=false){
					break;
				}
			}
		}
		
		if(rst!=true){
			exe_rst_status = 2;
			Logger.info("expData "+expData+"is not found !!" );
		}
	}

	public static void hoverTextbasedDynamicOperation4(ArrayList<String> list,String expData){
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		String element_3 = list.get(2);
		String element_4 = list.get(3);
		
		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		By by3 = returnByObject(element_3);
		By by4 = returnByObject(element_4);
		
		String actData = null ;
		
		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element "+listobj1.size());
		
		boolean rst =false;
		if(listobj1.size()>0){
			for(int i=0; i<listobj1.size(); i++){
				WebElement obj2 = listobj1.get(i);
				List<WebElement> listobj2 = obj2.findElements(by3);
				if(listobj2.size()>0){
					for(int loop2=0; loop2<listobj2.size(); loop2++){
						WebElement obj3 = listobj2.get(loop2);
						List<WebElement> listobj3 = obj3.findElements(by4);
						if(listobj3.size() > 0){
							for(int loop3=0; loop3<listobj3.size(); loop3++){
								WebElement obj4 = listobj3.get(loop3);
								actData = obj4.getText();
								Logger.info("actData "+actData);
								if(actData.equalsIgnoreCase(expData)){
									rst = true;
									Actions builder = new Actions(driver);
									builder.moveToElement(obj4).click().build().perform();
									exe_rst_status = 1;
									break;
								}
							}
						}
						if(rst!=false){
							break;
						}
					}
				}
				if(rst!=false){
					break;
				}
			}
		}
		
		if(rst!=true){
			exe_rst_status = 2;
			Logger.info("expData "+expData+" is not found !!" );
		}
	}

	
	public static void ClickTextbasedDynamicOperation5(ArrayList<String> list,String expData) {
		exe_rst_status = 2;
		String element_1 = list.get(0);
		String element_2 = list.get(1);
		String element_3 = list.get(2);
		String element_4 = list.get(3);
		String element_5 = list.get(4);

		By by1 = returnByObject(element_1);
		By by2 = returnByObject(element_2);
		By by3 = returnByObject(element_3);
		By by4 = returnByObject(element_4);
		By by5 = returnByObject(element_5);

		String actData = null;

		WebElement obj1 = driver.findElement(by1);
		List<WebElement> listobj1 = obj1.findElements(by2);
		Logger.info("Size of element " + listobj1.size());

		boolean rst = false;
		if (listobj1.size() > 0) {
			for (int i = 0; i < listobj1.size(); i++) {
				WebElement obj2 = listobj1.get(i);
				List<WebElement> listobj2 = obj2.findElements(by3);
				if (listobj2.size() > 0) {
					for (int loop2 = 0; loop2 < listobj2.size(); loop2++) {
						WebElement obj3 = listobj2.get(loop2);
						List<WebElement> listobj3 = obj3.findElements(by4);
						if (listobj3.size() > 0) {
							for (int loop3 = 0; loop3 < listobj3.size(); loop3++) {
								WebElement obj4 = listobj3.get(loop3);
								List<WebElement> listobj4 = obj4
										.findElements(by5);
								if (listobj4.size() > 0) {
									for (int loop4 = 0; loop4 < listobj4.size(); loop4++) {
										WebElement obj5 = listobj3.get(loop3);
										actData = obj5.getText();
										Logger.info("actData " + actData);
										if (actData.equalsIgnoreCase(expData)) {
											rst = true;
											obj5.click();
											exe_rst_status = 1;
											break;
										}
									}
								}
								if (rst != false) {
									break;
								}
							}
						}
						if (rst != false) {
							break;
						}
					}
				}
				if (rst != false) {
					break;
				}
			}
		}

		if (rst != true) {
			exe_rst_status = 2;
			Logger.info("expData " + expData + "is not found !!");
		}
	}
	
	//The parameters are Click and link for operationname
	public static void handlelinkOperation(String action, String operationname){
		
		int rownm = Driver.executing_row;
		//get the value for linkname from column 8 from excel i.e fieldvalue 2
		String lnkname = getValuefromExcel(Driver.TestData_Sheetpath, "Test Case",rownm, 8);
		//if we are clicking on a link, pass values like click and the name of the link from column 8
		if(action.equalsIgnoreCase("click")){
			clickLinkOperation(operationname,lnkname);
		}
		else if(action.equalsIgnoreCase("validate")){
			validateLinkOperation(lnkname);
		}
	}
	
	//this function is called for clicking on a link
	public static void clickLinkOperation(String oprnamekey, String lnkname){
		
		exe_rst_status = 2;
		
		By obj = null ;
		//if link is passed as paramter get the webelement object in obj
		if(oprnamekey.equalsIgnoreCase("link")){
			obj = By.linkText(lnkname);
		}
		//if partiallink is passed as parameter then get the webelement for the object in obj
		else if(oprnamekey.equalsIgnoreCase("partiallink")){
			obj = By.partialLinkText(lnkname);
		}
		
		boolean status = false;
		//find the webelement
		List<WebElement> ele = driver.findElements(obj);
		//if element found then click on the link usued the linkname
		if(ele.size()>0){
			if(oprnamekey.equalsIgnoreCase("link")){
				waitFortheElement(100, lnkname);
				driver.findElement(By.linkText(lnkname)).click();
				status = true;
				exe_rst_status = 1;
			}
			
			else if(oprnamekey.equalsIgnoreCase("partiallink")){
				waitFortheElement(100, lnkname);
				driver.findElement(By.partialLinkText(lnkname)).click();
				status = true;
				exe_rst_status = 1;
			}
		}
		//dynamic search lnk object and click it
		if(status != true){
			List<WebElement> ele1 = driver.findElements(By.tagName("a"));
			if(ele1.size()>0){
				for(int i=0; i<ele1.size(); i++){
					WebElement ele2 = ele1.get(i);
					List<WebElement> listele2 = ele2.findElements(By.tagName("span"));
					if(listele2.size()>0){
						for(int j=0; j<listele2.size();j++){
							WebElement ele3 = listele2.get(j);
							String actlnktxt = ele3.getText();
							if(actlnktxt.equalsIgnoreCase(lnkname)){
								//waitFortheElement(120, actlnktxt);
								ele3.click();
								Logger.info("link object "+lnkname+" found");
								status = true;
								exe_rst_status = 1;
								break;
							}
						}
					}
					if(status != false){
						break;
					}
				}
			}
		}
		
		//final link click status
		if(status!=true){
			Logger.info("Failed : link object "+lnkname+" not found !!");
			exe_rst_status = 2;
		}
		
	}
	
	public static void validateLinkOperation(String lnkname){
		
		exe_rst_status = 2;
		boolean status = false;
		
		List<WebElement> ele = driver.findElements(By.tagName("a"));
		if(ele.size()>0){
			for(int i=0; i<ele.size(); i++){
				WebElement ele2 = ele.get(i);
				String lnkdata = ele2.getText();
				if(lnkdata.equalsIgnoreCase(lnkname)){
					Logger.info("link match");
					exe_rst_status = 1;
					status = true;
					break;
				}
			}
		}
		//dynamic search lnk object and click it
		if(status != true){
			List<WebElement> ele1 = driver.findElements(By.tagName("a"));
			if(ele1.size()>0){
				for(int i=0; i<ele1.size(); i++){
					WebElement ele2 = ele1.get(i);
					List<WebElement> listele2 = ele2.findElements(By.tagName("span"));
					if(listele2.size()>0){
						for(int j=0; j<listele2.size();j++){
							WebElement ele3 = listele2.get(j);
							String actlnktxt = ele3.getText();
							if(actlnktxt.equalsIgnoreCase(lnkname)){
								Logger.info("link object "+lnkname+" found");
								status = true;
								exe_rst_status = 1;
								break;
							}
						}
					}
					if(status != false){
						break;
					}
				}
			}
		}
		
		//final link click status
		if(status!=true){
			Logger.info("Failed : link object "+lnkname+" not found !!");
			exe_rst_status = 2;
		}
		
	}
	
	public static int returnCurrentExecutionResultStatus(){
		return exe_rst_status;
	}
	//this method is for executing the different auto it operations and returning the status
	public static void executeAutoItOperation(String OperationName){
		exe_rst_status = 2;
		try{
			HashMap<String, Integer> map = new HashMap<String, Integer>();
			// Action mapping
			map.put("autoit_ff_fileUpload", 1);
			map.put("autoit_ff_openfile", 2);
			map.put("autoit_ie_openfile", 3);
			map.put("autoit_ie_savefile_2", 4);
			map.put("autoit", 5);
			map.put("new_autoit", 6);
			
			int data = 0;
			if (map.containsKey(OperationName)){
				 data = map.get(OperationName);
			}
			String msg = "operation perfomred!!";
			String initialized = "Autoit part initialized";
			String finalmsg = "autoit "+OperationName + " " + msg ;
			
			String autoitfilepath = Resourse_path.autoItpath+OperationName+".exe";
			Logger.info("autoitfilepath "+autoitfilepath);
			int rownm = Driver.executing_row;
			
			@SuppressWarnings("unused")
			Process pb ;
			
			switch(data){
			case 0:
				Logger.info(OperationName+" key does not match!!");
				exe_rst_status = 2;
				break;
			case 1:
				Logger.info(initialized);
				String path = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
				String filepath = Resourse_path.currPrjDirpath+File.separator+path;
				pb= new ProcessBuilder(autoitfilepath,filepath).start();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 2:
				Logger.info(initialized);
				String windowtitle = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
				pb= new ProcessBuilder(autoitfilepath,windowtitle).start();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 3:
				Logger.info(initialized);
				pb= new ProcessBuilder(autoitfilepath).start();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 4:
				Logger.info(initialized);
				pb= new ProcessBuilder(autoitfilepath).start();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;		
			case 5:
				Logger.info(initialized);
				String exefilename = getValuefromExcel(Driver.TestData_Sheetpath, Driver.DriverSheetname,rownm, 8);
				String exefilepath = Resourse_path.autoItpath+exefilename+".exe";
				pb= new ProcessBuilder(exefilepath).start();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 6:
				Logger.info(initialized);
				executeAutoItParameterBased();
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;		
			default:
				Logger.info(OperationName+" case is not mapped for this operation");
				exe_rst_status = 2;
				break;	
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	
	}
	
	public static void executeAutoItParameterBased(){
		try{
			String flnm = fn_Data("filename");
			String parameter = fn_Data("parameter");
			String [] arr = parameter.split(",");
			int par_len = arr.length;
			
			//defined null variables
			String var1 = null ;
			String var2 = null ;
			String var3 = null ;
			String var4 = null ;
			String var5 = null ;
			
			String fl = null ;
			if(flnm!=null){
				fl = flnm;
			}else{
				fl = "genericfunctions";
			}
			
			@SuppressWarnings("unused")
			Process pb ;
			String exefilepath = Resourse_path.autoItpath+fl+".exe";
			
			switch(par_len){
			case 1:
				var1 = arr[0];
				pb= new ProcessBuilder(exefilepath,var1).start();
				break;
				
			case 2:
				var1 = arr[0];
				var2 = arr[1];
				String path = fn_Data("path");
				if(path!=null){
					if(path.equalsIgnoreCase("default")){
						createDownloadFolderpath();
						var2 = Resourse_path.comp_downloadfoder+var2;
					}
				}
				pb= new ProcessBuilder(exefilepath,var1,var2).start();
				break;
				
			case 3:
				var1 = arr[0];
				var2 = arr[1];
				var3 = arr[2];
				pb= new ProcessBuilder(exefilepath,var1,var2,var3).start();
				break;
				
			case 4:
				var1 = arr[0];
				var2 = arr[1];
				var3 = arr[2];
				var4 = arr[3];
				pb= new ProcessBuilder(exefilepath,var1,var2,var3,var4).start();
				break;
				
			case 5:
				var1 = arr[0];
				var2 = arr[1];
				var3 = arr[2];
				var4 = arr[3];
				var5 = arr[4];
				pb= new ProcessBuilder(exefilepath,var1,var2,var3,var4,var5).start();
				break;	
				
			default:
				Logger.info(par_len+" case is not mapped for this operation");
				exe_rst_status = 2;
				break;	
			}
			resultmssg=var1+" Operation performed successfully!!";	
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//this function checks if the cell supplied as parameter has some value and returns true or false
	public static boolean validateCell(XSSFCell Cell_obj){
		boolean cellstatus = false;
		try{
			if(Cell_obj!=null){
		   		String data1 = Cell_obj.toString();
				Logger.info("col 6 data ; "+data1);
				char [] ch = data1.toCharArray();
				if(ch.length>0){
					cellstatus=true;
				}
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
		return cellstatus;
	}
	
	
	public static void fileUpload(String OR_ObjectName, String Path){
		exe_rst_status = 2;
		
		WebElement Object_name = GenericFunctions.Field_obj(OR_ObjectName);
		
		Object_name.sendKeys(Path);
		
		exe_rst_status = 1;
		
	}
	
	public static boolean isAlertPresent() {
		boolean presentFlag = false;
		try {
			@SuppressWarnings("unused")
			Alert alert = driver.switchTo().alert();
			// Alert present
			presentFlag = true;
		} catch (NoAlertPresentException ex) {
//			ex.printStackTrace();
			Logger.info("Alert is not present");
		}
		return presentFlag;

	}
	
	public static void staticwait(int time){
		try{
			Logger.info("Total static wait time "+time);
			for(int i=time; i>0; i--){
				Thread.sleep(1000);
				if(i%5==0){
					Logger.info("System waiting for the element.."+i+" sec to appear!!");
				}
			}
		}catch(Exception e){
			e.getMessage();
		}
	}
	
	  public static int returnFinalrow(XSSFSheet my_sheet){
		  int row = 0;
		  try{
			  int totalrow = my_sheet.getLastRowNum();
				Logger.info("totalrow "+totalrow+1);
				int introw = 0;
				for(int i=0; i<totalrow; i++){
					boolean rw = isRowEmptyInExcel(my_sheet.getRow(i));
					if(rw!=true){
						introw++;
					}
				}
				row = introw;
				Logger.info("Total final row "+row);
		  }catch(Exception e){
			  e.printStackTrace();
			  Logger.error(e);
		  }
		  return row;
	  }
	  
    //This function accepts the object name and object value as parameter and is for executing the input operation
	public static void executeInputOperation(String OR_ObjectName, String Value) {
		
		exe_rst_status = 2;
		//this function returns the webelement object for the object passed as parameter
		WebElement Object_name = GenericFunctions.Field_obj(OR_ObjectName);

		//gets the type of the object e.g checkbox, textbox etc
		String ObjectType = Object_name.getAttribute("Type");
		//gets the tagname for the object e.g. input, select etc
		String ObjectTag = Object_name.getTagName();

		Logger.info("ObjectTag " + ObjectTag + " and ObjectType "	+ ObjectType);
		// if object type is not found then mark as failed
		if (ObjectType == null) {
			ObjectType = Object_name.getAttribute("type");
			if(ObjectType==null){
//				Logger.info("Object Type is null");
				Logger.info("Object Type is null");
				ObjectType = "Cannot find";
//				exe_rst_status = 2;
			}
		}
		//if object tag cannot be found then mark as failed
		if (ObjectTag == null) {
//			Logger.info("Object Tag is null");
			Logger.info("Object Tag is null");
			ObjectTag = "Cannot find";
			exe_rst_status = 2;
		}
		
		//if object tag is select or option i.e dropdown, listbox or radio button etc
		if ((ObjectTag.equalsIgnoreCase("select")) || (ObjectTag.equalsIgnoreCase("option"))) {
			Logger.info("select combobox operation ");
			waitFortheElement(120, OR_ObjectName);
			SelectFromListBox(Object_name, Value);
		} 
		
		//this block is execute if we have to select a checkbox
		else if (ObjectType.equalsIgnoreCase("textarea")) {
			Logger.info("Enter data operation");
			waitFortheElement(120, OR_ObjectName);
			EnterText(Object_name, Value);
		}
		
		//this block is execute if we have to select a checkbox
		else if (ObjectType.equalsIgnoreCase("checkbox")) {
			Logger.info("select checkbox operation ");
			waitFortheElement(120, OR_ObjectName);
			SelectCheckbox(Object_name, Value);
		}
		//select checkbox
		else if(ObjectType.equalsIgnoreCase("checkbox") && (ObjectTag.equalsIgnoreCase("input"))){
			Logger.info("select checkbox operation");
			waitFortheElement(120, OR_ObjectName);
			SelectCheckbox(Object_name, Value);
		}
		//Text input field
		else if (ObjectType.equalsIgnoreCase("text") && ObjectTag.equalsIgnoreCase("input") || ObjectType.equalsIgnoreCase("text") || 
				ObjectType.equalsIgnoreCase("email") && ObjectTag.equalsIgnoreCase("input")) {
			Logger.info("Enter data operation");
			waitFortheElement(120, OR_ObjectName);
			EnterText(Object_name, Value);
		}
		//Text input field
		else if (ObjectType.equalsIgnoreCase("password") && ObjectTag.equalsIgnoreCase("input") || ObjectType.equalsIgnoreCase("password")) {
			Logger.info("Enter data operation");
			waitFortheElement(120, OR_ObjectName);
			EnterText(Object_name, Value);
		} 
		//select radio
		else if (ObjectType.equalsIgnoreCase("radio")) {
			Logger.info("Radio operation ");
			waitFortheElement(120, OR_ObjectName);
			RadioSelector(Object_name, Value);
		} 
		//select radio
		else if (ObjectType.equalsIgnoreCase("radio") && ObjectTag.equalsIgnoreCase("input")) {
			Logger.info("Radio operation ");
			waitFortheElement(120, OR_ObjectName);
			RadioSelector(Object_name, Value);
		} 
		//default message
		else {
			Logger.info(" Object type and tag has not been built yet");
			exe_rst_status = 2;
		}
		
	  }
	
	//This is for clicking on an object and accepts the object to be clicked and the fieldvalue if any
	public static void clickCondition(String OR_ObjectName, XSSFCell Cell_obj6){
		char [] ch = OR_ObjectName.toCharArray();
		if(ch.length>0){
			//if value present in object name
			Logger.info("Under single object handle operation!!");
			if(Cell_obj6!=null){
				//if value present in fieldvalue cell also 
	    		String col6 = Cell_obj6.toString();
				Logger.info("col 6 data : "+col6+" and col 5 data : "+OR_ObjectName);
				char [] chcol6 = col6.toCharArray();
				if(chcol6.length > 0){
					//if value present in fieldvalue cell then
					Logger.info("Click case operation !!");
					executeclickOps(OR_ObjectName,col6 );
				}else {
					Logger.info("Normal click operation !!");
					Click(OR_ObjectName);
				}
	    	}else {
				Logger.info("Normal click operation !!");
				Click(OR_ObjectName);
			}
		}else {
			Logger.info("Other click operation !!");
			//for clicking on links
			OtherClickOperation(Cell_obj6);
		}
	}
	
	/*
	public static void killProcess(String Name) {
		exe_rst_status = 2;
		try {
			WinNT winNT = (WinNT) Native.loadLibrary(WinNT.class,W32APIOptions.UNICODE_OPTIONS);
			WinNT.HANDLE snapshot = winNT.CreateToolhelp32Snapshot(Tlhelp32.TH32CS_SNAPPROCESS, new WinDef.DWORD(0));
			Tlhelp32.PROCESSENTRY32.ByReference processEntry = new Tlhelp32.PROCESSENTRY32.ByReference();
			String pname = Name+".exe";
			while (winNT.Process32Next(snapshot, processEntry)) {
				DWORD processid = processEntry.th32ProcessID;
				String processname = Native.toString(processEntry.szExeFile);
                // kiling process
				if (processname.equalsIgnoreCase(pname)) {
					Logger.info("processid " + processid+ " and processname " + processname);
					Runtime rt = Runtime.getRuntime();
					@SuppressWarnings("unused")
					Process proc = rt.exec("taskkill /F /IM " + processname);
					exe_rst_status = 1;
					Logger.info(processname+" Process killed");
				}
			}
			winNT.CloseHandle(snapshot);
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	}
	*/
	public static void validatekeyValueInExcel(String filepath, String sheetname, String key, String value){
		exe_rst_status = 2;
		try{
			if(filepath!=null){
				File file = new File(filepath);
				if(file.exists()){
					FileInputStream FIS = new FileInputStream(filepath);
					XSSFWorkbook Workbook_obj = new XSSFWorkbook(FIS);
					XSSFSheet sheet_obj = Workbook_obj.getSheet(sheetname);
					//total row return
					int finalrow = returnFinalrow(sheet_obj);
					boolean keys = false;
					boolean values =false;
					//logic
					for(int i=1; i<=finalrow; i++){
						XSSFRow row_obj = sheet_obj.getRow(i);
						if(row_obj!=null){
							XSSFCell cell_obj = row_obj.getCell(0);
							if(cell_obj!=null){
								String keynm = cell_obj.toString();
								if(keynm!=null && keynm.equalsIgnoreCase(key)){
									keys = true;
									int  cell_count = row_obj.getLastCellNum();
									for(int j=1; j<=cell_count; j++){
										XSSFCell cellvalue = row_obj.getCell(j);
										String keyvalue = cellvalue.toString();
										if(keyvalue!=null && keyvalue.equalsIgnoreCase(value)){
											values = true;
											exe_rst_status = 1;
											Logger.info("Key "+keynm +" and keyvalue "+keyvalue);
											Logger.info("Key Value matched!!");
											break;
										}
									}
									break;
								}
							}
						}
					}
					
					if(keys!=true){
						Logger.info("key "+key +" is not found !!");
					}
					if(values!=true){
						Logger.info("key "+key +" is found but value "+value+" is not matched");
					}
					
					Workbook_obj.close();
				}else{
					Logger.info("filepath does not exist!!");
				}
			}
			else {
				Logger.info("Contains null path");
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//this functions accepts the object and object value arrays as parameters and does the following checks
	//presence of the object, checks for expected value in a object
	public static void validateOperation(ArrayList<String> alName, ArrayList<String> alVal) {
		
		
		//get the size of fieldname and fieldvalue array lists
		int sA = alName.size();
		int sB = alVal.size();
		
		//if we have same number of fieldnames and fieldvalues
		if (sA==sB) {
			
			//Loop through all the fieldnames
			for (int i = 0; i < sA; i++) {
					//get the fieldname in a string
					String oName = alName.get(i);
					//get the fieldvalue in a string 
					String oVal = alVal.get(i);
					//if value equals "element check" call the ValidateElementPresence function and pass
					// the object name as parameter
					if (oVal.equalsIgnoreCase("element_check")) {
						validateElementPresence(oName);
					}
					//if object value has value as excel value
					else if (oVal.equalsIgnoreCase("excel_value")) {
						Logger.info("Under excel validation!!");
						//
						String keyname  = fn_Data("keyName");
						String keyvalue = fn_Data("keyValue");
						String filename = fn_Data("filename");
						String sheetname = fn_Data("sheetname");
						// return file path
						String filepath = returnfilePath(filename);
						validatekeyValueInExcel(filepath, sheetname, keyname, keyvalue);
					} 
					else {
						GenericFunctions.Validate(oName, oVal);
					}
			}
		}
	}
	
	
	public static void deleteFile(String filename){
		try{
			String dwnPath = Resourse_path.homepath+File.separator+"Downloads"+File.separator+filename;
			String tempPath = Resourse_path.homepath+File.separator+"AppData"+File.separator+"Local"+
			                  File.separator+"Temp"+File.separator+filename;
			String testdataPath = Resourse_path.currPrjDirpath+File.separator+"test-data"+File.separator+filename;
			String [] path = {dwnPath,tempPath,testdataPath};
			
			for(int i=0; i<path.length; i++){
				String pth = path[i];
				File file = new File(pth);
				if(file.exists()){
					file.delete();
					Logger.info("Deleted file : " +pth);
				}
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	public static String returnfilePath(String filename){
		String rpath = null;
		try{
			String dwnPath = Resourse_path.homepath+File.separator+"Downloads"+File.separator+filename;
			String tempPath = Resourse_path.homepath+File.separator+"AppData"+File.separator+"Local"+
			                  File.separator+"Temp"+File.separator+filename;
			String testdataPath = Resourse_path.currPrjDirpath+File.separator+"test-data"+File.separator+filename;
			
			String [] path = {dwnPath,tempPath,testdataPath};
			
			for(int i=0; i<path.length; i++){
				String pth = path[i];
				File file = new File(pth);
				if(file.exists()){
					rpath = pth;
					Logger.info("return filepath : " +pth);
					break;
				}
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
		return rpath;
	}
	
	  /**
     * Downloads a file from a URL
     * @param fileURL HTTP URL of the file to be downloaded
     * @param saveDir path of the directory to save the file
     * @throws IOException
     */
	public static void downloadFile() {
		WebElement Object_name = GenericFunctions.Field_obj(Driver.OR_ObjectName);
		String fileURL = Object_name.getAttribute("href");
		 String saveDir = Resourse_path.comp_downloadfoder ;
		try {
			
			int BUFFER_SIZE = 4096;
			URL url = new URL(fileURL);
			HttpURLConnection httpConn = (HttpURLConnection) url.openConnection();
			int responseCode = httpConn.getResponseCode();

			// always check HTTP response code first
			if (responseCode == HttpURLConnection.HTTP_OK) {
				String fileName = "";
				String disposition = httpConn.getHeaderField("Content-Disposition");
						
//				String contentType = httpConn.getContentType();
//				int contentLength = httpConn.getContentLength();

				if (disposition != null) {
					// extracts file name from header field
					int index = disposition.indexOf("filename=");
					if (index > 0) {
						fileName = disposition.substring(index + 10,disposition.length() - 1);
					}
				} else {
					// extracts file name from URL
					fileName = fileURL.substring(fileURL.lastIndexOf("/") + 1,fileURL.length());
				}
/*
				Logger.info("Content-Type = " + contentType);
				Logger.info("Content-Disposition = " + disposition);
				Logger.info("Content-Length = " + contentLength);
				Logger.info("fileName = " + fileName);
*/
				// opens input stream from the HTTP connection
				InputStream inputStream = httpConn.getInputStream();
				String saveFilePath = saveDir + File.separator + fileName;

				// opens an output stream to save into file
				FileOutputStream outputStream = new FileOutputStream(saveFilePath);
				int bytesRead = -1;
				byte[] buffer = new byte[BUFFER_SIZE];
				while ((bytesRead = inputStream.read(buffer)) != -1) {
					outputStream.write(buffer, 0, bytesRead);
				}

				outputStream.close();
				inputStream.close();
				Logger.info("File downloaded");
			} else {
				Logger.info("No file to download. Server replied HTTP code: "+ responseCode);
			}
			httpConn.disconnect();
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	public static void executevbscript(String vbsfile){
		try{
			String folder = Resourse_path.currPrjDirpath+"\\resources\\vbscript\\examples\\";
			String script = folder+vbsfile+".vbs";
			// search for real path:
			String executable = "wscript.exe";
			String cmdArr[] = { executable, script };
			Runtime.getRuntime().exec(cmdArr);
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//this function accepts the vbs file and the functionname as parameter
	public static void executevbscriptwithArguments(String vbsfile, String fname){
		exe_rst_status = 2;
		try{
			//name for the vbs file
			String script = vbsfile+".vbs";
			// search for real path:
			String executable = "wscript.exe";
			String cmdArr[] = { executable, script,fname};
			//executes the function in the vbs file using the windows shell
			Runtime.getRuntime().exec(cmdArr);
			exe_rst_status = 1;
		}catch(Exception e){
			exe_rst_status = 2;
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//this function is used for executing operations like vbscript, newautoit, presskey, switchtoiframe etc
	public static void executeOtherOperation(String OperationName){
		exe_rst_status = 2;
		resultmssg = null;
		try{
			HashMap<String, Integer> map = new HashMap<String, Integer>();
			// Action mapping
			map.put("vbscript", 1);
			map.put("new_autoit", 2);
			map.put("switchToIframe", 3);
			map.put("back_to_default_content", 4);
			map.put("presskey", 5);
			map.put("validte_filepresence", 6);
			
			int data = 0;
			if (map.containsKey(OperationName)){
				 data = map.get(OperationName);
			}
			
			String msg = "operation perfomred!!";
			String initialized = "vbscrpit part initialized";
			String finalmsg = OperationName + " " + msg ;
			
			String vbfolderpath = Resourse_path.vbfolderpath;
			String filename = vbfolderpath+fn_Data("filename");
			
			switch(data){
			case 0:
				Logger.info(OperationName+" key does not match!!");
				exe_rst_status = 2;
				break;
			case 1:
				String function_nm = fn_Data("function_name");
				Logger.info(initialized);
				File file = new File(filename+".vbs");
				Logger.info(filename+"exist");
				Logger.info("Executing vbscript..");
				if(file.exists()){
					//calls the function to execute vbscript, supplies the vs file and functionname
					//as parameter
					executevbscriptwithArguments(filename,function_nm);
				}else {
					Logger.info(filename+" does not exist!!");
					exe_rst_status = 2;
				}
				staticwait(5);
				Logger.info("vbscript executed!!");
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 2:
				Logger.info(msg);
				String del_fn = fn_Data("filename");
				//delte file before open if it exist..
				deleteFile(del_fn);
				staticwait(5);
				executeAutoItOperation(OperationName.toLowerCase());
				staticwait(12);
				exe_rst_status = 1;	
				break;
			case 3:
				switchToIframe(Driver.OR_ObjectName);
				break;
			case 4:
				returnToDefaultContent();
				break;	
			case 5:
				pressKey();
				break;	
			case 6:
				pressKey();
				break;		
			default:
				Logger.info(OperationName+" case is not mapped for this operation");
				exe_rst_status = 2;
				break;	
			}
		}catch(Exception e){
			exe_rst_status = 2;
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	public static void validateFilePresence(){
		try{
			exe_rst_status = 2;
			String filename = fn_Data("filename");
			if(filename!=null){
				//check in download folder
				String [] folderlist = {Resourse_path.comp_downloadfoder,Resourse_path.tempfolder,Resourse_path.prj_download_fld};
				boolean rstatus = false;
				
				for(int i=0; i<folderlist.length; i++){
					String folder = folderlist[i];
					File file = new File(folder+filename);
					if(file.exists()){
						Logger.info(filename+" File is present in :: "+folder);
						resultmssg = filename+" File is present";
						rstatus = true;
						exe_rst_status = 1;
					}
				}
				
				//Assigning result
				if(rstatus!=true){
					Logger.warn(filename+" is not present");
					resultmssg = filename+" File is not present";
				}
			}
			
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	//This function checks the presence of the Test Case sheet in the Workbook object for the Test Case file to be executed
	//It Logs a message incase the sheet is not found and also raises an exception
	public static void checkExcelSheetPresence(XSSFWorkbook Wbook_obj ){
		try{
			int shrs = 0;
		    // for each sheet in the workbook
			//validating sheet name 
            for (int sh = 0; sh < Wbook_obj.getNumberOfSheets(); sh++) {
            	String shnm = Wbook_obj.getSheetName(sh);
 //             Logger.info("Sheet name: " + shnm);
                Logger.info("Sheet name: " + shnm);
                if(shnm.equalsIgnoreCase("Test Case")){
                	//Logger.info(Driver.DriverSheetname+" sheetname found!!");
                	Logger.info(Driver.DriverSheetname+" sheetname found!!");
                	shrs=1;
                	break;
                }
            }
            
            if(shrs!=1){
            	//Logger.info(Driver.DriverSheetname+" sheetname not found!!");
            	Logger.info(Driver.DriverSheetname+" sheetname not found!!");
            }
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
	public static void driverCheck(){
		if(driver==null){
			Logger.info("Driver is null");
		}
	}
	
	public static void returnToDefaultContent(){
		exe_rst_status = 2;
		if(driver!=null){
			driver.switchTo().defaultContent();
		}
		exe_rst_status = 1;
	}
	
/** @author - ashish-choudhary
    @Function_Name -  switchToIframe()
    @Description - It switch the driver to frame
    @return WebElemet object
    @Created_Date - 26 Feb 2015
    @Modified_Date - 
 */
 	public static void switchToIframe(String OR_ObjectName) {
 		exe_rst_status = 2;
 		driverCheck();
 		HashMap<String, String> xml_property_storage = XMLReading(OR_ObjectName);

 		// Storing property name to arraylist
 		ArrayList<String> prpstorage = new ArrayList<String>();

 		for (Map.Entry<String, String> m : xml_property_storage.entrySet()) {
 			prpstorage.add(m.getKey());
 		}

 		ArrayList<String> localobjstorg = localObjStorage();

 		String obj = null;
 		boolean exitprcess = false;
 		boolean mapped = false;
 		for (int i = 0; i < prpstorage.size(); i++) {
 			obj = prpstorage.get(i);
 			// Logger.info("first "+obj);
 			for (int j = 0; j < localobjstorg.size(); j++) {
 				String obj2 = localobjstorg.get(j);
 				if (obj.equalsIgnoreCase(obj2)) {
 					mapped = true;
 					Logger.info(obj+" object successfully mapped!!");
 					String objvalue = xml_property_storage.get(obj.toLowerCase());
 					char[] ch = objvalue.toCharArray();
 					if (ch.length > 0) {
 						Logger.info(obj+" object contains property value !!");
 						exitprcess = true;
 						break;
 					}else{
 						Logger.info(obj+" object contains null property value, so moving to next object..");
 					}
 				}
 			}
 			if (exitprcess != false) {
 				break;
 			}
 		}

 		// object code validation
 		if (exitprcess != false) {

 			String prp_value = xml_property_storage.get(obj);
 			Logger.info("Property name - " + obj);
 			Logger.info("Property value - " + prp_value);

 			HashMap<String, Integer> swtich_prp_storage = switchPrpStorage();

 			// storing property name to hashmap
 			ArrayList<String> switchmapobject = new ArrayList<String>();
 			for (Map.Entry<String, Integer> m : swtich_prp_storage.entrySet()) {
 				switchmapobject.add(m.getKey());
 			}
 			// validating mapping with switch case
 			boolean switchprocess = false;
 			for (int i = 0; i < switchmapobject.size(); i++) {
 				String mpobject = switchmapobject.get(i);
 				if (obj.equalsIgnoreCase(mpobject)) {
 					switchprocess = true;
 					break;
 				}
 			}

 			// swtich case validation
 			if (switchprocess != false) {
 				int switchid = swtich_prp_storage.get(obj);
 				Logger.info(obj + " element is returned!!");
 				Logger.info("Finding "+obj+" element on webpage..");
 				// Assigning element to webelement
 				switch (switchid) {
 				case 1:
 					driver.switchTo().frame(driver.findElement(By.id(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 2:
 					driver.switchTo().frame(driver.findElement(By.xpath(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 3:
 					driver.switchTo().frame(driver.findElement(By.className(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 4:
 					driver.switchTo().frame(driver.findElement(By.name(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 5:
 					driver.switchTo().frame(driver.findElement(By.linkText(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 6:
 					driver.switchTo().frame(driver.findElement(By.tagName(prp_value)));
 					exe_rst_status = 1;
 					break;	
 				case 7:
 					driver.switchTo().frame(driver.findElement(By.cssSelector(prp_value)));
 					exe_rst_status = 1;
 					break;
 				case 8:
 					driver.switchTo().frame(driver.findElement(By.partialLinkText(prp_value)));
 					exe_rst_status = 1;
 					break;	
 				default:
 					Logger.info("Event is not mapped for this operation");
 					break;
 				}
 				Logger.info("Element found!!");
 			} else {
 				Logger.info(obj + " is not mapped in property local storage case!!");
 			}

 		} else {
 			if(xml_property_storage.size()>0){
 				if (mapped != true) {
 					Logger.info("Xml object is not mapped in local object storage!!");
 				}else {
 					if(exitprcess!=true){
 						Logger.info("Xml object is mapped but it does not contains object property value in xml file!!");
 					}
 				}
 			}else{
 				Logger.info("size of the element return by xmlreading is "+prpstorage.size());
 			}
 		}

 	}
 	
/** @author - ashish-choudhary
    @Function_Name -  switchToIframe()
    @Description - It switch the driver to frame
    @return WebElemet object
    @Created_Date - 26 Feb 2015
    @Modified_Date - 
 */
 	public static void switchToIframe() {
 		exe_rst_status = 2;
 		String indexvl = fn_Data("frameindex");
 		if(indexvl!=null){
 			char [] in_len = indexvl.toCharArray();
 			if(in_len.length>0){
 				int frameindex = Integer.parseInt(indexvl);
 	 	 		driver.switchTo().frame(frameindex);
 	 	 		exe_rst_status = 1;
 			}
 		}
 	}
 	
/** @author - brijesh-yadav
    @Function_Name -  
    @Description - 
    @return 
    @Created_Date - 
    @Modified_Date - 
 */
	public static void pressKey(){
		try{
			exe_rst_status = 2;	
			String key_nm = fn_Data("key");
			
			if(key_nm!=null){
				
				HashMap<String, Integer> map = new HashMap<String, Integer>();
				// Action mapping
				map.put("tab", 1);
				map.put("enter", 2);
				map.put("down", 3);
				
				// switching data operation
				int data = map.get(key_nm.toLowerCase());
				String msg = "operation perfomred!!";
				
				String stringloop = fn_Data("times");
				int loop = 1; 
				if(stringloop!=null){
					loop = Integer.parseInt(stringloop);
				}
				
				Robot robot = new Robot();
				switch(data){
				case 1:
					for(int i=1; i<=loop; i++){
						robot.keyPress(KeyEvent.VK_TAB);
						robot.keyRelease(KeyEvent.VK_TAB);
						Thread.sleep(1000);
					}
					Logger.info(key_nm + " " + msg);
					exe_rst_status = 1;	
					break;
				case 2:
					Logger.info(key_nm + " " + msg);
					exe_rst_status = 1;	
					break;
				case 3:
					break;
				case 4:
					Logger.info(key_nm + " " + msg);
					exe_rst_status = 1;	
					break;
				default:
					Logger.info(key_nm+" case is not mapped for this operation");
					break;
				}
			}
		}catch(Exception e){
			e.printStackTrace();
			Logger.error(e);
		}
	}
	
    public static void fillcommentinexcel(Workbook wbook_obj,Sheet wsheet_obj,Row row_obj,Cell cell, String exmsg){
		try {

			CreationHelper factory = wbook_obj.getCreationHelper();
			Drawing drawing = ((org.apache.poi.ss.usermodel.Sheet) wsheet_obj).createDrawingPatriarch();
			System.out.println("exmsg "+exmsg);
			String errormsg = exmsg;
			char[] ch = errormsg.toCharArray();
			int col = 3;
			int rowDimension = 5;
			System.out.println("ch "+ch.length);
			//providing area to comment box based on character length
			if (ch.length > 100 && ch.length <= 500) {
				col += 3;
				rowDimension += 3;
			} else if (ch.length > 500 && ch.length <= 1000) {
				col += 5;
				rowDimension += 6;
			} else if (ch.length > 1000 && ch.length <= 2000) {
				col += 10;
				rowDimension += 18;
			} else if (ch.length > 2000) {
				col += 20;
				rowDimension += 30;
				System.out.println("third");
			}
			// When the comment box is visible, have it show in a 1x3 space
			ClientAnchor anchor = factory.createClientAnchor();
			anchor.setCol1(cell.getColumnIndex());
			anchor.setCol2(cell.getColumnIndex() + col);
			anchor.setRow1(row_obj.getRowNum());
			anchor.setRow2(row_obj.getRowNum() + rowDimension);
			// Create the comment and set the text+author
			Comment comment = drawing.createCellComment(anchor);
			RichTextString str = factory.createRichTextString(errormsg);
			comment.setString(str);
			// Assign the comment to the cell
			cell.setCellComment(comment);
			
		} catch (Exception e) {
			e.printStackTrace();
			Logger.error(e);
		}
    }
    
    //store stacktrace to string
    public static String getStackTrace(Exception e) {
        final Writer result = new StringWriter();
        final PrintWriter printWriter = new PrintWriter(result);
        e.printStackTrace(printWriter);
        return result.toString();
      }
    
    public static void createDownloadFolderpath(){
    	try{
    		String path = Resourse_path.date_time_down_folder+Resourse_path.running_path;
    		Logger.info(path);
    		File file = new File(path);
    		if(!file.exists()){
    			file.mkdirs();
    			Logger.info("Download fodler path created!!");
    		}
    		//Assigning download folder path
    		Resourse_path.comp_downloadfoder = path;
    	}catch(Exception e){
    		e.printStackTrace();
			Logger.error(e);
    	}
    }
    
    public static String getDate(){
    	String datevalue = null;
    	DateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd");
		// get current date time with Date()
		Date date = new Date();
		datevalue = dateFormat.format(date);
		return datevalue;
    }
    
    public static String returntime(){
    	String time = null;
		DateFormat datetime = new SimpleDateFormat("HH_mm_ss");
		// get current date time with Calendar()
		Calendar cal = Calendar.getInstance();
		time = datetime.format(cal.getTime());
		return time;
    }
    
    public static void createDateTimeFolder(String foldernm){
    	try{
    		String date = getDate();
    		String time = returntime();
    		String down_date_time_folder = Resourse_path.currPrjDirpath+File.separator+foldernm+File.separator
    				      +date+File.separator+"time_"+time+File.separator;
    		
    		File file = new File(down_date_time_folder);
    		if(!file.exists()){
    			file.mkdirs();
    			Logger.info("Download fodler path created!!");
    		}
    		//assigning path to static string
    		Resourse_path.date_time_down_folder=down_date_time_folder;
    		Logger.info("date_time_down_folder "+Resourse_path.date_time_down_folder);
    	}catch(Exception e){
    		e.printStackTrace();
			Logger.error(e);
    	}
    }
    
    //store values to static till complete execution
    public static void session_setvalue_single(){
    	String flag_v = fn_Data("flag_a");
    	if(flag_v!=null){
    		Resourse_path.flag_a=flag_v;
    	}
    }
    
    public static void vm_mouseclick(String OperationName){
    	try{
    		HashMap<String, Integer> map = new HashMap<String, Integer>();
			// Action mapping
			map.put("vbscript", 1);
			map.put("new_autoit", 2);
			
			int data = 0;
			if (map.containsKey(OperationName)){
				 data = map.get(OperationName);
			}
			
			String msg = "operation perfomred!!";
			String finalmsg = OperationName + " " + msg ;
			String init = OperationName + " initialized!!" ;
			
			switch(data){
			case 0:
				Logger.info(OperationName+" key does not match!!");
				exe_rst_status = 2;
				break;
			case 1:
				Logger.info(init);
				//code goes here
				
				Logger.info(finalmsg);
				exe_rst_status = 1;
				break;
			case 2:
				Logger.info(init);
				//code goes here
				
				Logger.info(finalmsg);
				exe_rst_status = 1;	
				break;
			default:
				Logger.info(OperationName+" case is not mapped for this operation");
				exe_rst_status = 2;
				break;	
			}
    	}catch(Exception e){
    		e.printStackTrace();
    		Logger.error(e);
    	}
    }
    
    
    public static void csvDataEmployee(String input, String Output) throws Exception{
		
		File inputFile = new File(input);
		File outputFile = new File(Output);
		BufferedReader br;
		BufferedWriter writer;
		br = new BufferedReader(new FileReader(inputFile));
		writer = new BufferedWriter(new FileWriter(outputFile));
	        
	        ArrayList<String> lines = new ArrayList<String>();
	        
	        
	        String line = null;
	        while ((line = br.readLine()) != null) {
	            lines.add(line);
	        }
	        
	        int p = lines.size();
	        
	        long reCount;
	        reCount=0;
	        
        	
        	String[] arr1,arr2,arrTemp;
        	
	        arr1 = lines.get(0).split(",");
	        arrTemp = lines.get(p-1).split(",");
	        
	        for (int k=0; k<=arr1.length-1; k++) {
	        	
	        	String tempHead = arr1[k].toLowerCase();
	        	
	        	if (tempHead.equals("hremployeeid")) {
	        		
	        		String temp = arrTemp[k];
	        		
	        		String[] newArrTemp = temp.split("_");
	        		
	        		reCount = Integer.parseInt(newArrTemp[2]);
	        	}
	        }
	        
	        
	        String[] newArrTemp;
    		
    		String NewReCount="";
    		String newString="";
    		String temp="";
    		
    		String newHeader;
    		newHeader = "";
        	
        	for (int xz=0; xz<=arr1.length-1; xz++) {
        		
        		newHeader = newHeader+arr1[xz]+",";
        		
        	}
    		
    		writer.write(newHeader + System.getProperty("line.separator"));
	        
	        for (int i=1; i<p; i++) {
	        	
	        	arr2 = lines.get(i).split(",");
	        	
	        	reCount = reCount+1;
	        	
	        	for (int j=0; j<=arr1.length-1; j++) {
	        		
	        		String tempHead = arr1[j].toLowerCase();
		        	
		        	if (tempHead.equals("hremployeeid")) {
		        		
		        		temp = arr2[j];
		        		
		        		newArrTemp = temp.split("_");
		        		
		        		NewReCount = String.valueOf(reCount);
		        		
		        		newString = newArrTemp[0]+"_"+newArrTemp[1]+"_"+NewReCount;
		        		
		        	}
		        	
		        	if (tempHead.equals("hremployeeid") || tempHead.equals("nationalidno") || tempHead.equals("fullname") || tempHead.equals("username")) {
		        		
		        		arr2[j] = newString;
		        	}
	        	}
	        	String newRecord;
	        	newRecord = "";
	        	
	        	for (int z=0; z<=arr2.length-1; z++) {
	        		newRecord = newRecord+arr2[z]+",";
	        	}
	        	writer.write(newRecord + System.getProperty("line.separator"));
	        }
	        br.close();	
	        writer.close(); 
	        if(!inputFile.delete()){
	        	throw new Exception("exception in ouputfile generation-deleting.");
	        }
	        if(!outputFile.renameTo(inputFile)){
	        	throw new Exception("exception in ouputfile generation-renaming.");
	        	}
	 }

public static void csvDataDependent(String input, String Output) throws Exception{
	
	File inputFile = new File(input);
	File outputFile = new File(Output);
	BufferedReader br;
	BufferedWriter writer;
	br = new BufferedReader(new FileReader(inputFile));
	writer = new BufferedWriter(new FileWriter(outputFile));
	
        ArrayList<String> lines = new ArrayList<String>();
        
        
        
        String line = null;
        while ((line = br.readLine()) != null) {
            lines.add(line);
        }
        
        int p = lines.size();
        
        long reCount;
        reCount=0;
        
    	
    	String[] arr1,arr2,arrTemp;
    	
        arr1 = lines.get(0).split(",");
        arrTemp = lines.get(p-1).split(",");
        
        for (int k=0; k<=arr1.length-1; k++) {
        	
        	String tempHead = arr1[k].toLowerCase();
        	
        	if (tempHead.equals("hremployeeid")) {
        		
        		String temp = arrTemp[k];
        		
        		String[] newArrTemp = temp.split("_");
        		
        		reCount = Integer.parseInt(newArrTemp[2]);
        	}
        }
        
        
        String[] newArrTemp;
        String[] newArrTemp2;
		
		String NewReCount = "";
		String newString = "";
		String newString2 = "";
		String temp = "";
		String temp2 = "";
		
		String newHeader;
		newHeader = "";
		
		String tempPrev;
		tempPrev = "";
		
    	
    	for (int xz=0; xz<=arr1.length-1; xz++) {
    		
    		newHeader = newHeader+arr1[xz]+",";
    		
    	}
		
		writer.write(newHeader + System.getProperty("line.separator"));
        
        for (int i=1; i<p; i++) {
        	
        	arr2 = lines.get(i).split(",");
        	
        	for (int j=0; j<=arr1.length-1; j++) {
        		
        		String tempHead = arr1[j].toLowerCase();
	        	
	        	if (tempHead.equals("hremployeeid")) {
	        		
	        		
	        		if (!tempPrev.equals(arr2[j])) {
	        			
	        			reCount = reCount+1;
	        			
	        			tempPrev = arr2[j];
	        			
	        			temp = arr2[j];
		        		
		        		newArrTemp = temp.split("_");
		        		
		        		NewReCount = String.valueOf(reCount);
		        		
		        		newString = newArrTemp[0]+"_"+newArrTemp[1]+"_"+NewReCount;
	        			
	        		}
	        		
	        		arr2[j] = newString;
	        	}
	        	
	        	
	        	if (tempHead.equals("nationalidno") || tempHead.equals("fullname")) {
	        		
	        		temp2 = arr2[j];
	        		
	        		newArrTemp2 = temp2.split("_");
	        		
	        		newString2 = newString+"_"+newArrTemp2[3];
	        		
	        		arr2[j] = newString2; 
	        		
	        	}
        	}
        	
        	String newRecord;
        	newRecord = "";
        	
        	for (int z=0; z<=arr2.length-1; z++) {
        		
        		newRecord = newRecord+arr2[z]+",";
        		
        	}
        	
        	writer.write(newRecord + System.getProperty("line.separator"));
        	
        }
        
        
        br.close();	
        writer.close(); 
        if(!inputFile.delete()){
        	throw new Exception("exception in ouputfile generation-deleting.");
        }
        if(!outputFile.renameTo(inputFile)){
        	throw new Exception("exception in ouputfile generation-renaming.");
        	}
        	
	
 }


public static void copyCSVEmp(String input1, String Output) throws Exception{
	
	File inputFile1 = new File(input1);
	
	BufferedReader br1;

	br1 = new BufferedReader(new FileReader(inputFile1));
			
	ArrayList<String> lines1 = new ArrayList<String>();
	
    String line1 = null;
    while ((line1 = br1.readLine()) != null) {
        lines1.add(line1);
    }
	
    File myFile = new File(Output);
    FileInputStream fis = new FileInputStream(myFile);

    XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
    
    int a = myWorkBook.getNumberOfSheets();
	
	
	int flagE=0;
	
	
	for (int i = 0; i < a; i++) {
		
		XSSFSheet Sheet1 = myWorkBook.getSheetAt(i);
		
		if (Sheet1.getSheetName().equals("EmpData")) {
			flagE=1;
			break;
		}	
	
	}
	
	
	int temp1=0;
	
	if (flagE==1) {
		temp1 = myWorkBook.getSheetIndex("EmpData");
		myWorkBook.removeSheetAt(temp1);
		myWorkBook.createSheet("EmpData");
	} else if (flagE==0) {
		myWorkBook.createSheet("EmpData");
	}
	
	XSSFSheet Employee = myWorkBook.getSheet("EmpData");
	
    HashMap<Integer, String[]> dataEmp = new HashMap<Integer, String[]>();
    
    int p1 = lines1.size();
    
    String arr1;
	String[] newArr1;
	
    for (int j = 0; j < p1; j++) {
    	
    	arr1 = lines1.get(j);
    	newArr1 = arr1.split(",");
    	
    	dataEmp.put(j+1, newArr1);
    	
    }
    
	Set<Integer> newRowsE = dataEmp.keySet();
	
    int rownumE = Employee.getLastRowNum();         
 
    for (Integer keyE : newRowsE) {
     
        Row rowE = Employee.createRow(rownumE++);
        Object [] objArrE = dataEmp.get(keyE);
        int cellnumE = 0;
        for (Object objE : objArrE) {
            Cell cellE = rowE.createCell(cellnumE++);
            if (objE instanceof String) {
            	
                cellE.setCellValue((String) objE);
                
            } else if (objE instanceof Boolean) {
            	
                cellE.setCellValue((Boolean) objE);
                
            } else if (objE instanceof Date) {
            	
                cellE.setCellValue((Date) objE);
                
            } else if (objE instanceof Double) {
            	
                cellE.setCellValue((Double) objE);
                
            }
        }
    }

   
    myWorkBook.getCreationHelper().createFormulaEvaluator().evaluateAll();
    
    FileOutputStream os = new FileOutputStream(Output);
    myWorkBook.write(os);
    
	br1.close();
	myWorkBook.close();

}

public static void copyCSVDep(String input2, String Output) throws Exception{
	
	File inputFile2 = new File(input2);

	BufferedReader br2;

	br2 = new BufferedReader(new FileReader(inputFile2));
	
	ArrayList<String> lines2 = new ArrayList<String>();
    
    String line2 = null;
    while ((line2 = br2.readLine()) != null) {
        lines2.add(line2);
    }
    
    File myFile = new File(Output);
    FileInputStream fis = new FileInputStream(myFile);
	
    XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
    
    int a = myWorkBook.getNumberOfSheets();
	
	int flagD;
	flagD=0;
			
	for (int i = 0; i < a; i++) {
		
		XSSFSheet Sheet1 = myWorkBook.getSheetAt(i);
		
		if (Sheet1.getSheetName().equals("DepData")) {
			flagD=1;
		}
		
	}
	
	int temp2;
	temp2=0;
	
	if (flagD==1) {
		temp2 = myWorkBook.getSheetIndex("DepData");
		myWorkBook.removeSheetAt(temp2);
		myWorkBook.createSheet("DepData");
	} else if (flagD==0) {
		myWorkBook.createSheet("DepData");
	}
	
	XSSFSheet Dependent = myWorkBook.getSheet("DepData");
    
    HashMap<Integer, String[]> dataDep = new HashMap<Integer, String[]>();
    
    int p2 = lines2.size();
	
	String arr2;
	String[] newArr2;
	
    for (int k = 0; k < p2; k++) {
    	
    	arr2 = lines2.get(k);
    	newArr2 = arr2.split(",");
    	
    	dataDep.put(k+1, newArr2);
    	
    }
    
    Set<Integer> newRowsD = dataDep.keySet();
	
    int rownumD = Dependent.getLastRowNum();         
 
    for (Integer keyD : newRowsD) {
     
        Row rowD = Dependent.createRow(rownumD++);
        Object [] objArrD = dataDep.get(keyD);
        int cellnumD = 0;
        for (Object objD : objArrD) {
            Cell cellD = rowD.createCell(cellnumD++);
            if (objD instanceof String) {
            	
                cellD.setCellValue((String) objD);
                
            } else if (objD instanceof Boolean) {
            	
                cellD.setCellValue((Boolean) objD);
                
            } else if (objD instanceof Date) {
            	
                cellD.setCellValue((Date) objD);
                
            } else if (objD instanceof Double) {
            	
                cellD.setCellValue((Double) objD);
                
            }
        }
    }
    
   myWorkBook.getCreationHelper().createFormulaEvaluator().evaluateAll();

    FileOutputStream os = new FileOutputStream(Output);
    myWorkBook.write(os);
    
	br2.close();
	myWorkBook.close();

}



public static void csvOperations(String SheetPath1) {
		
		String input_e = null;
		String input_d = null;
		String input_h = null;
		String TestCase;
		
		
		if (Driver.Empcsvfname!= "NA")
		{
			input_e = Resourse_path.csv_path+"AutomationSG/"+Driver.Empcsvfname;
		}
		
		if (Driver.Depcsvfname!= "NA")
		{
			input_d = Resourse_path.csv_path+"AutomationSG/"+Driver.Depcsvfname;
		}
		if (Driver.Histcsvfname!= "NA")
		{
			input_h = Resourse_path.csv_path+"AutomationSG/"+Driver.Histcsvfname;
		}
	
		TestCase = SheetPath1;
		
		try {
//			csvDataEmployee(input_e, Output_e);
//			csvDataDependent(input_d, Output_d);
			BenefitAsiaFunction cv = new BenefitAsiaFunction();
			cv.generateCode(input_e,input_d, input_h);
			if (input_e!=null)
			{
				copyCSVEmp(input_e, TestCase);
			}
			if(input_d!=null)
			{
				copyCSVDep(input_d, TestCase);
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}
	
	
/*	String input_e, Output_e, input_d, Output_d, TestCase;
	input_e = Resourse_path.csv_path+"/AutomationSG/Employee_1.csv";
	Output_e = Resourse_path.csv_path+"/AutomationSG/output_Employee_1.csv";

	input_d = Resourse_path.csv_path+"/AutomationSG/Dependent_1.csv";
	Output_d = Resourse_path.csv_path+"/AutomationSG/output_Dependent_1.csv";
	
	//TestData_Sheetpath = Resourse_path.TestData_Sheetpath+DriverSheetname+".xlsx";

	TestCase = SheetPath1;
	
	//TestCase = TestData_Sheetpath;


	try {
		csvDataEmployee(input_e, Output_e);
	} catch (Exception e1) {
		// TODO Auto-generated catch block
		e1.printStackTrace();
	}

	try {
		csvDataDependent(input_d, Output_d);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}

	try {
		copyCSVEmp(input_e, TestCase);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	try {
		copyCSVDep(input_d, TestCase);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	*/
	




/*

public static void validate_Excel_Data( String ResultFolderPath, ArrayList<String> FieldName, ArrayList<String> FieldValue) 
{
	
	
	XSSFWorkbook  workbookobj = null;
	String File_name = null;
		
	exe_rst_status = 2;
	try {
		File_name = FieldValue.get(0);
		String Sheet_Path = ResultFolderPath+"/"+File_name+".xlsx";
		FileInputStream xlfile = new FileInputStream(new File(Sheet_Path));
	     
	    //Get the workbook instance for XLS file 
	    workbookobj = new XSSFWorkbook(xlfile);
	 
	    //Get first sheet from the workbook
	    XSSFSheet wksheet = workbookobj.getSheetAt(0);
		
        //Row count operation
		int totalRows = wksheet.getLastRowNum()+1;
		int introw = 0;
		for (int f = 1; f < totalRows; f++) 
		{
			boolean rw_rs = GenericFunctions.isRowEmptyInExcel(wksheet.getRow(f));
			if (rw_rs != true)
			{
				introw++;
			}
		}
		
        //final row count
		int finaltotalRows = introw;
		
		String PK_FieldName = FieldName.get(1);
		String PK_FieldValue = FieldValue.get(1);
		int PK_Row = 0;
		int HeaderRow = 0;
		int PK_Field_ColNum1 = 0;
		Row rowctr = wksheet.getRow(HeaderRow);
		int TotalColumns = rowctr.getLastCellNum();
		
		
		for (int colnum = 0; colnum < TotalColumns;colnum++)
		{
		 String colvalue = rowctr.getCell(colnum).getStringCellValue();
		 if (PK_FieldName.equals(colvalue))
		 {
			 PK_Field_ColNum1 = colnum;
			 System.out.println("PK Field Header Found");
			 break;
			 
		 }
		}
		
		for (int i = 1; i<= finaltotalRows; i++ )
			
		{
			XSSFRow row_obj = wksheet.getRow(i);
			XSSFCell cell_obj = row_obj.getCell(PK_Field_ColNum1);
			String PK_FieldVal = cell_obj.getStringCellValue();
			if (PK_FieldVal.equalsIgnoreCase(PK_FieldValue))
			{
				PK_Row = i;
				 System.out.println("PK Field Value Found");
				break;
			}
		}
		Row rowctr1 = wksheet.getRow(HeaderRow);
		for(int k=2; k<FieldName.size();k++)
		{//loop will execute till size of arraylist.
			    for(int m = 0; m<TotalColumns; m++)
			   {
				   String colval = rowctr1.getCell(m).getStringCellValue();//get value of column headers
				   if (FieldName.get(k).equalsIgnoreCase(colval))//if column header matches
				   {
					   XSSFRow row_obj1 = wksheet.getRow(PK_Row);//set counter to PK Row
					   //Row rowctr2 = wksheet.getRow(PK_Row);
					   String ExpFldVal = FieldValue.get(k);//get expected value from fieldvalue arraylist
					   XSSFCell cell_obj1 = row_obj1.getCell(m);
						String Act_FieldVal = cell_obj1.getStringCellValue();//get actual value from sheet
					   if (ExpFldVal.equalsIgnoreCase(Act_FieldVal)) //value matches in sheet and array list
					   {
						   System.out.println("success");
					   }
						else
						{
							System.out.println("failure");
						   
						  }
				   }
					   
				   
			   }

		}

	}
		
	
	catch (Exception e) {
		e.printStackTrace();
	}
	}
}
	
*/


//This function fetches the value from the object specified and stores in the treemap declared in Driver 
public static void GetVal(String OR_ObjectName, String val){
	exe_rst_status = 2;
	resultmssg = null;
	//This finds the object and returns the webelemnt object
	WebElement Object_name=GenericFunctions.Field_obj(OR_ObjectName);
	driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
	//get the actual value from the object using getText
	String RunTimeValue = Object_name.getText();
	Logger.info("Run time value fetched is " + RunTimeValue);
	String Keyvalue = Driver.Testcasename + "|" + Driver.StepNumber + "|" + OR_ObjectName ;
	Logger.info("Key value created is " + Keyvalue);
	if (RunTimeValue != "" || Keyvalue != "")
	{
		//converts the value to lowercase
		Driver.GVmap.put(Keyvalue, RunTimeValue);
		exe_rst_status = 1;
		resultmssg = "Value fetched from : "+OR_ObjectName+" and stored in : "+ Keyvalue;
	} 
	else
	{
		resultmssg = "Value cannot be fetched from : "+OR_ObjectName+" and not stored in : "+ Keyvalue;;
		exe_rst_status = 2;
	}	
}




//this functions accepts the fieldname and fieldvalue array lists as parameters and calls the Getval function
	public static void GetValue(ArrayList<String> FName, ArrayList<String> FVal)
	{
		
		
		//get the size of fieldname and fieldvalue array lists
		int sA = FName.size();
		int sB = FVal.size();
		
		//if we have same number of fieldnames and fieldvalues
		if (sA==sB)
		{
			
			//Loop through all the fieldnames
			for (int i = 0; i < sA; i++)
			{
					//get the fieldname in a string
					String oName = FName.get(i);
					//get the fieldvalue in a string 
					String oVal = FVal.get(i);
					GenericFunctions.GetVal(oName, oVal);
			}
		}
	}
	
public static void validate_Excel_Data( String ResultFolderPath, ArrayList<String> FieldName, ArrayList<String> FieldValue) throws IOException 
{
       XSSFWorkbook  workbookobj = null;
       String File_name = null;
       String colheader_flag = "false";
       String colvalue_flag = "false";
       DataFormatter objDefaultFormat = new DataFormatter();
       FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbookobj);
       exe_rst_status = 1;
       try {
              File_name = FieldValue.get(0);
              ResultFolderPath="C:\\BA2Automation\\Results";
              String Sheet_Path = ResultFolderPath+"/"+File_name+".xlsx";
              FileInputStream xlfile = new FileInputStream(new File(Sheet_Path));
           
              
              
           //Get the workbook instance for XLS file 
           workbookobj = new XSSFWorkbook(xlfile);
           try{
           int shrs = 0;
                     
           //Get first sheet from the workbook
           XSSFSheet wksheet = workbookobj.getSheetAt(0);
           
              
        //Row count operation
              int totalRows = wksheet.getLastRowNum()+1;
              int introw = 0;
              for (int f = 1; f < totalRows; f++) 
              {
                    boolean rw_rs = GenericFunctions.isRowEmptyInExcel(wksheet.getRow(f));
                    if (rw_rs != true)
                    {
                           introw++;
                    }
              }
              
        //get the total rows in the excel sheet
              int finaltotalRows = introw;
              //the primary key field is mentioned in the first element of the arraylist
              String PK_FieldName = FieldName.get(1);
              boolean field = false;
              //if primary keu field like employee id is not mentioned in excel
              //else not mentioned if pk field and pk value not enetred in excel code should exit and fail
              if(PK_FieldName.equals(null))
              {
                    Logger.info("Primary key field not Entered");
                    field = true;
              }
              //the primary key value is mentioned in te first element of the fieldvalue array list
              String PK_FieldValue = FieldValue.get(1);
              
              if(PK_FieldName.equals(null)){
                    Logger.info("Primary key value not Entered");
                    field = true;
              }
              
              
              int PK_Row = 4;//the row from where the values start
              int HeaderRow = 3;//the row where the column headers are mentioned
              int PK_Field_ColNum1 = 0;
              Row rowctr = wksheet.getRow(HeaderRow);//set the row counter to the header row
              int TotalColumns = rowctr.getLastCellNum(); // get the total number of columns
              
              //loop through all the columns in the header row and find the primary key field column header and 
              //store it's column number in the PK_Field_ColNum1 variable
              //else not mentioned the code should exit with fatal error if pk header column in not found
              for (int colnum = 0; colnum < TotalColumns ; colnum++)
              {
              String colvalue = rowctr.getCell(colnum).getStringCellValue();
              if (PK_FieldName.equals(colvalue))
              {
                    PK_Field_ColNum1 = colnum;
                     System.out.println("PK Field Header Found");
                    break;
                    
               }
              }
              
              //loop starting from the row from where values start to the end of excel
              for (int i = PK_Row; i<= finaltotalRows; i++ )
                    
              {
            	  //set the counter on first row from where values start
                    XSSFRow row_obj = wksheet.getRow(i);
                    
                    XSSFCell cell_obj = row_obj.getCell(PK_Field_ColNum1);
                    String PK_FieldVal = cell_obj.getStringCellValue();
                    if (PK_FieldVal.equalsIgnoreCase(PK_FieldValue))
                    {
                           PK_Row = i;
                           System.out.println("PK Field Value Found");
                           break;
                    }
              }
              Row rowctr1 = wksheet.getRow(HeaderRow);
              String ExpFldVal = null;
              String Act_FieldVal= null;
              
              for(int k=2; k<FieldName.size();k++)
              {//loop will execute till size of arraylist.
                        for(int m = 0; m<TotalColumns; m++)
                       {
                              String colval = rowctr1.getCell(m).getStringCellValue();//get value of column headers
                              if (FieldName.get(k).equalsIgnoreCase(colval))//if column header matches
                              {
                                   colheader_flag= "true";  
                            	  Logger.info("Field name " + FieldName.get(k) + "is found in the header row");
                            	  XSSFRow row_obj1 = wksheet.getRow(PK_Row);//set counter to PK Row
                                     //Row rowctr2 = wksheet.getRow(PK_Row);
                                    ExpFldVal = FieldValue.get(k);//get expected value from fieldvalue arraylist
                                     XSSFCell cell_obj1 = row_obj1.getCell(m);
                                     objFormulaEvaluator.evaluate(cell_obj1); 
                                     Act_FieldVal = objDefaultFormat.formatCellValue(cell_obj1,objFormulaEvaluator);
                                    // Act_FieldVal = cell_obj1.getStringCellValue();//get actual value from sheet
                                     if (ExpFldVal.equalsIgnoreCase(Act_FieldVal)) //value matches in sheet and array list
                                     {
                                    	 colvalue_flag = "true";  
                                    	 Logger.info("Expected Value" + ExpFldVal +  "Actual Value" +  Act_FieldVal);
                                    	 break;
                                     }
                                   
                              }
                             
                           if(m == TotalColumns-1)
                           {
                        	   if (colheader_flag.equalsIgnoreCase("false")|| colvalue_flag.equalsIgnoreCase("false"))
                        	   {
                        		   Logger.info("Expected Value" + ExpFldVal +  "Actual Value" +  Act_FieldVal);
                        		   exe_rst_status = 2;
                        	   }
                           }
                              
                       }

              }

       }
              
       
       catch (Exception e)
       {
              e.printStackTrace();
       }
      
       }
       finally
       {
    	   workbookobj.close(); 
       }
       }


public static FirefoxProfile FirefoxDriverProfile() throws Exception {
	FirefoxProfile profile = new FirefoxProfile();
	profile.setPreference("browser.download.folderList", 2);
	profile.setPreference("browser.download.manager.showWhenStarting", false);
	profile.setPreference("browser.download.dir", "C:\\Talent_Automation_Framework_Backup\\Project_Talent\\Results");
	profile.setPreference("browser.helperApps.neverAsk.openFile", false);
	profile.setPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
	profile.setPreference("browser.helperApps.alwaysAsk.force", false);
	profile.setPreference("browser.download.manager.alertOnEXEOpen", false);
	profile.setPreference("browser.download.manager.focusWhenStarting", false);
	profile.setPreference("browser.download.manager.useWindow", false);
	profile.setPreference("browser.download.manager.showAlertOnComplete", false);
	profile.setPreference("browser.download.manager.closeWhenDone", false);
	return profile;
}



public static void validate_action_list(String action_name){
	exe_rst_status = 2;
	String msg = "operation perfomred!!";
	String finalmsg = action_name + " " + msg ;
	
	switch(action_name){
	
	case "contains_text":
		WebElement Object_name = GenericFunctions.Field_obj(Driver.OR_ObjectName);
		String textdata = Object_name.getText();
		if(textdata.length() > 1){
			exe_rst_status = 1;
			resultmssg = "Contains the data and total length of character is :"+textdata.length();
		}else {
			resultmssg = "No text data is available";
		}
		Logger.info(finalmsg);
		break;
	case "group_data":
		String object_list = Driver.OR_ObjectName;
		String [] object_arr_list = object_list.split(",");
		
		break;
	default:
		Logger.info(action_name+" case is not mapped for this operation");
		break;	
	}
	
}


}







