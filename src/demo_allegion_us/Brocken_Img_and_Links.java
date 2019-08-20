package demo_allegion_us;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Brocken_Img_and_Links {

	public static void main(String[] args) throws IOException, InterruptedException {
		System.setProperty("webdriver.gecko.driver", "./Drivers\\geckodriver.exe");
		WebDriver  driver = new FirefoxDriver();
		
		
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		
		 File src=new File("./Input_File\\client_site.xlsx");

		   // load file
		   FileInputStream fis=new FileInputStream(src);
		 
		   // Load workbook
		   XSSFWorkbook wb=new XSSFWorkbook(fis);
		   
		   // Load sheet- Here we are loading first sheetonly
		      XSSFSheet sh1= wb.getSheetAt(0);
		 
		  // getRow() specify which row we want to read.

		  // and getCell() specify which column to read.
		  // getStringCellValue() specify that we are reading String data.
		   for(int k=0; k<=sh1.getLastRowNum(); k++)
		   {
		    	  String value = sh1.getRow(k).getCell(0).getStringCellValue();
		    	  System.out.println("website in test --->" + value);
		    	  
		    	  Thread.sleep(3000);   	
		driver.get(value);
		Thread.sleep(3000);
		
		List <WebElement> linkslist = driver.findElements(By.tagName("a"));
			 linkslist.addAll(driver.findElements(By.tagName("img")));
		
		System.out.println("size of full links and images --->" +linkslist.size());
		
		
		List <WebElement> activelinks = new ArrayList <WebElement> ();
		
		
		for(int i=0; i<linkslist.size(); i++)
		{
			if (linkslist.get(i).getAttribute("href")!= null && (!linkslist.get(i).getAttribute("href").contains("javascript")))
			{
				activelinks.add(linkslist.get(i));
			}
				
		}
		System.out.println("size of active links and images --->" + activelinks.size());
		
	
	    Sheet sheet2 = wb.createSheet("brokenlinks"+ k);
	    sheet2.createRow(0).createCell(0).setCellValue("website in test --->" + value);
	    sheet2.createRow(1).createCell(0).setCellValue("size of full links and images --->" +linkslist.size());
	    sheet2.createRow(2).createCell(0).setCellValue("size of active links and images --->" + activelinks.size());
		for(int j=0; j<activelinks.size(); j++)
		{
			HttpURLConnection connection = (HttpURLConnection) new URL(activelinks.get(j).getAttribute("href")).openConnection(); 
			
			connection.connect();
			String response = connection.getResponseMessage();
			connection.disconnect();
			Row head = sheet2.createRow(4);
			head.createCell(0).setCellValue("URL");
			head.createCell(1).setCellValue("RESPONSE");
			head.createCell(2).setCellValue("RESPONSE CODE");
			
			Row row = sheet2.createRow(j+5);
			Cell cell1 = row.createCell(0);
			Cell cell2 = row.createCell(1);
			Cell cell3 = row.createCell(2);
		
			cell1.setCellValue(activelinks.get(j).getAttribute("href"));
			cell2.setCellValue(response);
			cell3.setCellValue(connection.getResponseCode());
			System.out.println(activelinks.get(j).getAttribute("href") + "--->" + response+connection.getResponseCode());
				   
		}
		
		linkslist.clear();
	    activelinks.clear();
		
}
		FileOutputStream fileOut = new FileOutputStream("./Output_File\\Broken_links_imgs.xlsx");
		   wb.write(fileOut);
		   fileOut.close();
		  
 		
	}
}


