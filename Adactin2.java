package org.adactin2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Adactin2 {
	
	
	public static void main(String[] args) throws IOException, InterruptedException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		for (int i = 0; i < 6; i++) {
			
		driver.get("https://adactinhotelapp.com/");
		driver.manage().window().maximize();
		
		File file = new File("C:\\Users\\ELCOT\\eclipse-workspace\\DayTwoClass\\src\\test\\resources\\ExcelFiles\\Sheet1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("TestData");
		
		WebElement username = driver.findElement(By.id("username"));
		WebElement password = driver.findElement(By.id("password"));
		username.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
		password.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
		
		WebElement btnlogin = driver.findElement(By.id("login"));
		btnlogin.click();
		
		WebElement location = driver.findElement(By.id("location"));
		
		Select select = new Select(location);
		select.selectByVisibleText(sheet.getRow(i).getCell(2).getStringCellValue());
		
		WebElement hotels = driver.findElement(By.id("hotels"));
		
		Select select1 = new Select(hotels);
		select1.selectByVisibleText(sheet.getRow(i).getCell(3).getStringCellValue());
		
		WebElement roomtype = driver.findElement(By.id("room_type"));
		
		Select select2 = new Select(roomtype);
		select2.selectByVisibleText(sheet.getRow(i).getCell(4).getStringCellValue());
		
		WebElement rooms = driver.findElement(By.id("room_nos"));
		
		Select select3 = new Select(rooms);
		select3.selectByVisibleText(sheet.getRow(i).getCell(5).getStringCellValue());
		
		WebElement datein = driver.findElement(By.id("datepick_in"));
		Date din = (Date)sheet.getRow(i).getCell(6).getDateCellValue();
		datein.sendKeys(String.valueOf(din));
		
		WebElement dateout = driver.findElement(By.id("datepick_out"));
		Date dout = (Date)sheet.getRow(i).getCell(7).getDateCellValue();
		dateout.sendKeys(String.valueOf(dout));
		
		WebElement adultnos = driver.findElement(By.id("adult_room"));
		
		Select select4 = new Select(adultnos);
		select4.selectByVisibleText(sheet.getRow(i).getCell(8).getStringCellValue());
		
		WebElement childnos = driver.findElement(By.id("child_room"));
		
		Select select5 = new Select(childnos);
		select5.selectByVisibleText(sheet.getRow(i).getCell(9).getStringCellValue());
		
		WebElement submit = driver.findElement(By.id("Submit"));
		submit.click();
		
		WebElement btn = driver.findElement(By.id("radiobutton_0"));
		btn.click();
		
		WebElement btn2 = driver.findElement(By.id("continue"));
		btn2.click();
		
		WebElement firstelement = driver.findElement(By.id("first_name"));
		firstelement.sendKeys(sheet.getRow(i).getCell(10).getStringCellValue());
		
		WebElement lastname = driver.findElement(By.id("last_name"));
		lastname.sendKeys(sheet.getRow(i).getCell(11).getStringCellValue());
		
		WebElement address = driver.findElement(By.id("address"));
		address.sendKeys(sheet.getRow(i).getCell(12).getStringCellValue());
		
		WebElement cc = driver.findElement(By.id("cc_num"));
		long l = (long)sheet.getRow(i).getCell(13).getNumericCellValue();
		cc.sendKeys(String.valueOf(l));
		
		WebElement cctype = driver.findElement(By.id("cc_type"));
		
		Select select6 = new Select(cctype);
		select6.selectByVisibleText(sheet.getRow(i).getCell(14).getStringCellValue());
		
		WebElement exp = driver.findElement(By.id("cc_exp_month"));
		
		Select select7 = new Select(exp);
		select7.selectByVisibleText(sheet.getRow(i).getCell(15).getStringCellValue());
		
		WebElement expyear = driver.findElement(By.id("cc_exp_year"));
		long li = (long)sheet.getRow(i).getCell(16).getNumericCellValue();
		expyear.sendKeys(String.valueOf(li));
		
		
		WebElement ccv = driver.findElement(By.id("cc_cvv"));
		int k = (int)sheet.getRow(i).getCell(17).getNumericCellValue();
		ccv.sendKeys(String.valueOf(k));
		
		WebElement btnbook = driver.findElement(By.id("book_now"));
		btnbook.click();
		
		
		Thread.sleep(5000);
		WebElement order1 = driver.findElement(By.xpath("//input[@name='order_no']"));
		String attribute = order1.getAttribute("value");
		System.out.println(attribute);
		Row row = sheet.getRow(i);
		Cell createCell = row.createCell(18);
		createCell.setCellValue(attribute);
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);
		
				
	}

}
}