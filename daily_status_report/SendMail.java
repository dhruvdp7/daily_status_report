package com.qait.daily_status_report.daily_status_report;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class SendMail {
	static WebDriver driver;

	public static String to_data, cc_data, subject_data, description_data;

	public void readExcel(String filePath, String fileName, String sheetName) throws IOException {

		File file = new File(filePath + "/" + fileName);

		FileInputStream inputStream = new FileInputStream(file);

		Workbook myWorkbook = null;

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		myWorkbook = new HSSFWorkbook(inputStream);

		Sheet mySheet = myWorkbook.getSheet(sheetName);

		int rowCount = mySheet.getLastRowNum() - mySheet.getFirstRowNum();

		Row row = mySheet.getRow(1);

		to_data = mySheet.getRow(0).getCell(1).getStringCellValue();
		cc_data = mySheet.getRow(1).getCell(1).getStringCellValue();
		subject_data = mySheet.getRow(2).getCell(1).getStringCellValue();
		description_data = mySheet.getRow(3).getCell(1).getStringCellValue();
	}

	public static void main(String... strings) throws IOException {

		SendMail objExcelFile = new SendMail();

		String filePath = System.getProperty("user.dir")
				+ "/src/main/java/com/qait/daily_status_report/daily_status_report";

		objExcelFile.readExcel(filePath, "email.xls", "EmailData");

		System.setProperty("webdriver.chrome.driver", "/home/qainfotech/chromedriver");
		driver = (WebDriver) new ChromeDriver();
		WebDriverWait wait=new WebDriverWait(driver, 20);
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS) ;
		driver.get("https://webmail.qainfotech.com");
		 driver.findElement(By.id("username")).sendKeys("dhruvpande@qainfotech.com");
		 driver.findElement(By.id("password")).sendKeys("QA_webmail.1122");
		 driver.findElement(By.className("ZLoginButton")).click();
		 WebElement new_message = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("zb__NEW_MENU_title")));
		 new_message.click();
		 
		 
		 try{
			 Thread.sleep(2000);
			 }
			 catch(Exception e){
				 System.out.println(e);
			 }
		 driver.findElement(By.id("zv__COMPOSE-1_to_control")).sendKeys(to_data);
		
		 driver.findElement(By.id("zv__COMPOSE-1_cc_control")).sendKeys(cc_data);
		
		 driver.findElement(By.id("zv__COMPOSE-1_subject_control")).sendKeys(subject_data);
		 driver.switchTo().frame(0);
		 driver.findElement(By.id("tinymce")).sendKeys(description_data);
		 driver.switchTo().parentFrame();
		 driver.findElement(By.id("zb__COMPOSE-1__SEND_title")).click();
	}

}