package driverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunctions.FuntiontionLibrary;
import utilities.XLfile_Util;

public class DriverScript {
	WebDriver driver;
	String inputpath="./FileInput/Data_Controller.xlsx";
    String outputpath="./FileOutput/HybridResults.xlsx";
	ExtentReports reports;
	ExtentTest logger;
	String TCSheet="MasterTestCases";

	public void start_test()throws Throwable {
		
   //step-10
		String Module_Status="";
		String Module_New="";
	//Step-1	//Create object for xLfile_util
		XLfile_Util xl=new XLfile_Util(inputpath);
   //Step-2	//iterate all rows
		for (int i = 1; i<=xl.rowcount(TCSheet); i++) {
			
   //Step-3		//condition for "Master test cases" sheet
			if (xl.getcelldata(TCSheet, i, 2).equalsIgnoreCase("Y"))
			{
  //Step-5  //read module cell From TCSheet
				String TCModule=xl.getcelldata(TCSheet, i, 1);
				
  //Step-14	 //Define path of h.t.m.l
				reports=new ExtentReports("./target/Reports/"+TCModule+FuntiontionLibrary.generateDate()+".html");
				logger=reports.startTest(TCModule);
  //Step-6	     //Iterate the TCModule all rows
				for (int j =1; j <=xl.rowcount(TCModule); j++) {
 //Step-7	    //Read the cells in TCModule
					String Description=xl.getcelldata(TCModule, j, 0);
					String objtype=xl.getcelldata(TCModule, j, 1);
					String Ltype=xl.getcelldata(TCModule, j, 2);
					String Lvalue=xl.getcelldata(TCModule, j, 3);
					String Testdata=xl.getcelldata(TCModule, j, 4);
	    //Step-8     use try & catch
					try {
						
        //step-9    call all methods in "FunctionLibrary" class				
						if (objtype.equalsIgnoreCase("startBrowser")) {
						driver=FuntiontionLibrary.startBrowser();
						logger.log(LogStatus.INFO, Description);
						}
						
						if (objtype.equalsIgnoreCase("openUrl")) {
							FuntiontionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
							}
//---------------------------ApplicationLogin(module) sheet methods-----------------------------------------
						if (objtype.equalsIgnoreCase("waitForElement")) {
						    FuntiontionLibrary.waitForElement(Ltype, Lvalue, Testdata);
							logger.log(LogStatus.INFO, Description);
							}
						if (objtype.equalsIgnoreCase("typeAction")) {
							FuntiontionLibrary.typeAction(Ltype, Lvalue, Testdata);
							logger.log(LogStatus.INFO, Description);
							}
						if (objtype.equalsIgnoreCase("clickAction")) {
							FuntiontionLibrary.clickAction(Ltype, Lvalue);
							logger.log(LogStatus.INFO, Description);
							}
						if (objtype.equalsIgnoreCase("validateTitle")) {
							FuntiontionLibrary.validateTitle(Testdata);
							logger.log(LogStatus.INFO, Description);
						}
						if (objtype.equalsIgnoreCase("closeBrowser")) {
							FuntiontionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
//---------Step-16-----------------StockItems(module) sheet methods-------------------------------------------							
						
						if (objtype.equalsIgnoreCase("dropDownAction")) {
							FuntiontionLibrary.dropDownAction(Ltype, Lvalue, Testdata);
							logger.log(LogStatus.INFO, Description);
						}
						if (objtype.equalsIgnoreCase("captureStock")) {
							FuntiontionLibrary.captureStock(Ltype, Lvalue);
							logger.log(LogStatus.INFO, Description);
						}
						if (objtype.equalsIgnoreCase("stockTable")) {
							FuntiontionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
//-------Step-17---------------------Suppliers(module) sheet methods----------------------------------------
						if (objtype.equalsIgnoreCase("captureSupplier")) {
							FuntiontionLibrary.captureSupplier(Ltype, Lvalue);
							logger.log(LogStatus.INFO, Description);
						}
							if (objtype.equalsIgnoreCase("supplierTable")) {
								FuntiontionLibrary.supplierTable();
								logger.log(LogStatus.INFO, Description);}
 //-------Step-18-------------------Customers(module) sheet methods--------------------------------------------
							
							if (objtype.equalsIgnoreCase("captureCustomer")) {
								FuntiontionLibrary.captureCustomer(Ltype, Lvalue);
								logger.log(LogStatus.INFO, Description);
							}
							if (objtype.equalsIgnoreCase("customerTable")) {
								FuntiontionLibrary.customerTable();
								logger.log(LogStatus.INFO, Description);
							}
//step-11  //write as "PASS" into status cell in to TCModule sheet
						xl.setcelldata(TCModule, j, 5, "Pass", outputpath);
						logger.log(LogStatus.PASS , Description);
						Module_Status="true";
						
	//step_8	
					} catch (Exception e) 
					{
					System.out.println(e.getMessage());
					
 //step-12	 //write as "FAIL" into status cell in to TCModule sheet
					xl.setcelldata(TCModule, j, 5, "Fail", outputpath);
					logger.log(LogStatus.FAIL, Description);
					Module_New="false";
		//Screenshot for Failed test cases
					File screnshot=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(screnshot, new File("./target/Screenshot/"+Description+FuntiontionLibrary.generateDate()+".jpg"));
					}
					
	// step-13	 write  as "PASS" status in Master test case sheet 		
					if (Module_Status.equalsIgnoreCase("True")) {
						xl.setcelldata(TCSheet, i, 3, "PASS", outputpath);
					}
	// Step-15  //For Stop generated h.t.m.l Reports
					reports.endTest(logger);
					reports.flush();
					}
	// step-13	 write  as "FAIL" status in Master test case sheet
				if (Module_New.equalsIgnoreCase("False")) {
					xl.setcelldata(TCSheet, i, 3, "FAIL", outputpath);
				}
			}
			
   //Step-4  //Write as Blocked in status cell for test case flag to 'N'
			else {
			xl.setcelldata(TCSheet, i, 3, "Blocked", outputpath);
			}
		}
	}
}
