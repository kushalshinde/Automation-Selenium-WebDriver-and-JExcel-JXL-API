package com.project.TestSuiteBase;

import java.util.*;
import java.util.concurrent.TimeUnit;
import java.io.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


public class SuiteBase {
	
	public static WebDriver driver=null;
	public boolean BrowseralreadyLoaded=false;
	public static File exlFile = null;
	public static WritableWorkbook writableWorkbook = null;
	public static WritableSheet writableSheet = null;
	public boolean Excelcreated= false;
	
	public static Map<Integer, ArrayList<Integer>> map = new TreeMap<Integer, ArrayList<Integer>>();
	
	public void loadWebBrowser() throws Exception{
			System.setProperty("webdriver.chrome.driver","C:\\Users\\kusha\\eclipse-workspace\\SeleniumProject\\JarFiles\\chromedriver.exe");
			if(!BrowseralreadyLoaded && !Excelcreated){		
				ChromeOptions options = new ChromeOptions(); 
				options.addArguments("disable-infobars"); 
				
				driver = new ChromeDriver(options);
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				
				driver.get("https://www.cs.rutgers.edu/people/directory.php?type=faculty");
				waitFor(2000);
				
				exlFile = new File("C:\\Users\\kusha\\eclipse-workspace\\SeleniumProject\\src\\com\\project\\ExcelFiles\\rutgers.xls");
				writableWorkbook = Workbook.createWorkbook(exlFile);
				writableSheet = writableWorkbook.createSheet("Sheet1", 0);
				
				Label nameOfUniversity = new Label(0,0,"RUTEGERS CS FACULTY");
				writableSheet.addCell(nameOfUniversity);
				
				Label name = new Label(0,1,"NAME");
				writableSheet.addCell(name);
				
				Label phoneNumber = new Label(1,1,"PHONE NUMBER");
				writableSheet.addCell(phoneNumber);
				
				
				Label Email = new Label(2,1,"EMAIL");
				writableSheet.addCell(Email);
				
				Label HyperLinkAddress = new Label(3,1,"LINK");
				writableSheet.addCell(HyperLinkAddress);
				
				
				BrowseralreadyLoaded=true;
				Excelcreated= true;
				
			}
		}
	
	
	public void closeWebBrowser(){
		driver.close();
		try {
				Thread.sleep(2000);
		        Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe");
			} catch (InterruptedException | IOException e1) {
					e1.printStackTrace();
			}
		
	}
	
	public void putintomap() {
		for(int i=2; i<67; i=i+5) {
			 List<Integer> list= new ArrayList<Integer>();

			 if(i==2) {
				 for(int j=1; j<=5; j=j+2) {
					 //System.out.println(i + " "+j);
					 list.add(j);
					 if(j==5) {
						 map.put(i, new ArrayList<Integer>(list));
					 }
				 }
			 }else {
				 for(int k=1; k<=4; k+=k) {
					 //System.out.println(i + " "+k);
					 list.add(k);
					 if(k==4) {
						 map.put(i, new ArrayList<Integer>(list));
					 }
				 }
			 }
		 }
	}
	
	public void waitFor(int durationInMilliSeconds) {
        try {
            Thread.sleep(durationInMilliSeconds);
        } catch (InterruptedException e) {
            e.printStackTrace(); 
        }
    }
	
	public boolean isElementPresent(By by) {
        try {
            driver.findElement(by);
            return true;
        } catch (NoSuchElementException e) {
            return false;
        }
    }
}
