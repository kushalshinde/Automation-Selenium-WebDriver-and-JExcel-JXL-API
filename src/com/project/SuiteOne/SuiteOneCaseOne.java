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
		String name1 = "";
		String phone1 = "";
		String hyperlink1 = "";
		String email = "";
		int m = 2;
		for(Entry<Integer, ArrayList<Integer>> e : map.entrySet()) {
			Integer key = e.getKey();		
			for(Integer s : e.getValue()) {
				name1 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[1]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getText();
				phone1 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[1]/div/div["+key+"]/div["+s+"]/div/div[3]")).getText();
				hyperlink1 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[1]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getAttribute("href");
				email = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[1]/div/div["+key+"]/div["+s+"]/div/div[4]/a/span")).getText();
					
				Label n1 = new Label(0,m,name1);
				Label p1 = new Label(1,m, phone1);
				Label hl1 = new Label(3,m,hyperlink1);
				Label e1 = new Label(2, m, email);
				m++;
				writableSheet.addCell(n1);
				writableSheet.addCell(p1);
				writableSheet.addCell(hl1);
				writableSheet.addCell(e1);
						
				name1="";
				phone1 = "";
				hyperlink1="";
				email = "";
			}
		}
	}
	
	@Test
	public void techingfaculty() throws Exception{
		
		Label techingfaculty = new Label(0,43,"TEACHING FACULTY");
		writableSheet.addCell(techingfaculty);
		String name2 = "";
		String phone2 = "";
		String hyperlink2 = "";
		String email2 = "";
		
		int n = 44;
		
		for(Entry<Integer, ArrayList<Integer>> e : map.entrySet()) {
			Integer key = e.getKey();		
			if(key>=2 && key<=12) {
				for(Integer s : e.getValue()) {
					name2 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[2]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getText();
					phone2 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[2]/div/div["+key+"]/div["+s+"]/div/div[3]")).getText();
					hyperlink2 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[2]/div/div["+key+"]/div["+s+"]/div/div[2]/a")).getAttribute("href");
					email2 = driver.findElement(By.xpath("html/body/div[3]/div[1]/div[2]/div/div["+key+"]/div["+s+"]/div/div[4]/a/span")).getText();
						
					Label n2 = new Label(0,n,name2);
					Label p2 = new Label(1,n, phone2);
					Label hl2 = new Label(3,n,hyperlink2);
					Label e2 = new Label(2, n, email2);
					n++;
					writableSheet.addCell(n2);
					writableSheet.addCell(p2);
					writableSheet.addCell(hl2);
					writableSheet.addCell(e2);
							
					name2="";
					phone2 = "";
					hyperlink2="";
					email2 = "";
				}
			}
		}
	}
	
	@AfterTest
	public void tearDown() throws Exception {
		writableWorkbook.write();
        writableWorkbook.close();
        closeWebBrowser();
	}
}
