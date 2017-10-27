package com.project.SuiteOne;


import java.util.ArrayList;
import java.util.Map.Entry;

import org.openqa.selenium.By;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.project.TestSuiteBase.SuiteBase;

import jxl.write.Label;
import jxl.write.WriteException;



public class SuiteOneCaseOne extends SuiteBase {
	
	@BeforeTest
	public void runFirst() throws Exception{
		loadWebBrowser();
		putintomap();
	}
	
	
	@Test
	public void newrun() throws Exception, WriteException {
		
		int divelement=1;
		int m = 2;
		for(Entry<Integer, ArrayList<Integer>> e : map.entrySet()) {
			Integer key = e.getKey();		
			for(Integer s : e.getValue()) 
				writeintoExcel(key, s, divelement, m++);
		}
	}
	
	@Test
	public void techingfaculty() throws Exception{
		
		Label techingfaculty = new Label(0,43,"TEACHING FACULTY");
		writableSheet.addCell(techingfaculty);
		int divelement=2;
		int m = 44;
		
		for(Entry<Integer, ArrayList<Integer>> e : map.entrySet()) {
			Integer key = e.getKey();		
			if(key>=2 && key<=12) {
				for(Integer s : e.getValue())
					writeintoExcel(key, s, divelement, m++);
			}
		}
	}
	
	public void writeintoExcel(Integer key, Integer s, int divelement, int m) throws Exception, WriteException {
		String name = "";
		String phone = "";
		String hyperlink = "";
		String email = "";
		name = driver.findElement(By.xpath("html/body/div[3]/div[1]/div["+divelement+"]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getText();
		phone = driver.findElement(By.xpath("html/body/div[3]/div[1]/div["+divelement+"]/div/div["+key+"]/div["+s+"]/div/div[3]")).getText();
		hyperlink = driver.findElement(By.xpath("html/body/div[3]/div[1]/div["+divelement+"]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getAttribute("href");
		email = driver.findElement(By.xpath("html/body/div[3]/div[1]/div["+divelement+"]/div/div["+key+"]/div["+s+"]/div/div[4]/a/span")).getText();
			
		Label n1 = new Label(0,m,name);
		Label p1 = new Label(1,m, phone);
		Label hl1 = new Label(3,m,hyperlink);
		Label e1 = new Label(2, m, email);
		writableSheet.addCell(n1);
		writableSheet.addCell(p1);
		writableSheet.addCell(hl1);
		writableSheet.addCell(e1);
				
		}
	
	@AfterTest
	public void tearDown() throws Exception {
		writableWorkbook.write();
        writableWorkbook.close();
        closeWebBrowser();
	}
}
