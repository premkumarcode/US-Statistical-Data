import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class US_StatisticalData_NorthAmerican_MountainRanges {
	public static class Globals {
		   public static int ExcelRowCount = 1;
	}
	public static void main(String[] args) throws InterruptedException, IOException {
		//Variable Declaration
		String BrowserDriverPath = ".\\BrowserDriver\\";
		String Browser_to_launch="Chrome";
		String WebsiteLaunched="https://www.peakbagger.com/range.aspx?rid=1";
		WebDriver BrwsrLaunch=LaunchBrowser(Browser_to_launch,BrowserDriverPath);
		String Parent_Path="//td[contains(text(),'(Parent)')]/preceding-sibling::td//a";
			
		BrwsrLaunch.get(WebsiteLaunched);
		BrwsrLaunch.manage().window().maximize();
		JavascriptExecutor js = (JavascriptExecutor) BrwsrLaunch;
		List<String> L2=Populate_LevelChild(2,BrwsrLaunch,"","","","","","");
		System.out.println("Level 2 list " + L2);
		for(int L2c=0;L2c<L2.size();L2c++) {
			List<String> L3=Populate_LevelChild(3,BrwsrLaunch,L2.get(L2c).toString(),L2.get(L2c).toString(),"","","","");
			System.out.println("Level 2 : "+L2.get(L2c));
			System.out.println("Level 3 list " + L3);
			for(int L3c=0;L3c<L3.size();L3c++) {
				List<String> L4=Populate_LevelChild(4,BrwsrLaunch,L3.get(L3c).toString(),L2.get(L2c).toString(),L3.get(L3c).toString(),"","","");
				System.out.println("Level 2 : "+L2.get(L2c) + " Level 3 : "+L3.get(L3c));
				System.out.println("Level 4 list " + L4);
				for(int L4c=0;L4c<L4.size();L4c++) {
					List<String> L5=Populate_LevelChild(5,BrwsrLaunch,L4.get(L4c).toString(),L2.get(L2c).toString(),L3.get(L3c).toString(),L4.get(L4c).toString(),"","");
					System.out.println("Level 2 : "+L2.get(L2c) + " Level 3 : "+L3.get(L3c) + " Level 4 : "+L4.get(L4c));
					System.out.println("Level 5 list " + L5);
					for(int L5c=0;L5c<L5.size();L5c++) {
						System.out.println("L5 list size is : " + L5.size() + " L5c counter - " + L5c);
						List<String> L6=Populate_LevelChild(6,BrwsrLaunch,L5.get(L5c).toString(),L2.get(L2c).toString(),L3.get(L3c).toString(),L4.get(L4c).toString(),L5.get(L5c).toString(),"");
						BrwsrLaunch.findElement(By.xpath(Parent_Path)).click();
						System.out.println("Level 2 : "+L2.get(L2c) + " Level 3 : "+L3.get(L3c) + " Level 4 : "+L4.get(L4c) + " Level 5 : "+L5.get(L5c));
						System.out.println("Level 6 list " + L6);
					}
					BrwsrLaunch.findElement(By.xpath(Parent_Path)).click();
				}
				BrwsrLaunch.findElement(By.xpath(Parent_Path)).click();
			}
			BrwsrLaunch.findElement(By.xpath(Parent_Path)).click();
		}
		BrwsrLaunch.findElement(By.xpath(Parent_Path)).click();	
	}
	
	static List<String> Populate_LevelChild(int Level,WebDriver WB,String Lnktxt,String E_L2,String E_L3,String E_L4,String E_L5,String E_L6) throws InterruptedException, IOException{
		String Parent_Path="";
		String Child_Xpath="";
		String Nav_level_down_Xpath="";
		String CC="";
		if(Lnktxt.contains("'")) {
			int idx=Lnktxt.indexOf("'");
			Lnktxt=Lnktxt.substring(idx+1);
			System.out.println(Lnktxt);
		}
		if(!Lnktxt.equals("")) {
			Nav_level_down_Xpath="//td[contains(text(),'Level "+ (Level-1) +" (Child)')]/preceding-sibling::td//a[contains(text(),'" + Lnktxt + "')]";
			WB.findElement(By.xpath(Nav_level_down_Xpath)).click();
		}
		
		
		Child_Xpath="//td[contains(text(),'Level "+ (Level) +" (Child)')]/preceding-sibling::td//a";

		Parent_Path="//td[contains(text(),'(Parent)')]/preceding-sibling::td//a";

		
		List<WebElement> LevelChild = WB.findElements(By.xpath(Child_Xpath));
		List<String> LvlRngeLst = new ArrayList<String>();
		
		if(!LevelChild.isEmpty()) {
			for(int cntr=0;cntr<LevelChild.size();cntr++)
			{
				List<WebElement> TWE = WB.findElements(By.xpath(Child_Xpath));
				String Ranges=TWE.get(cntr).getText();
				Thread.sleep(2000);
				TWE.get(cntr).click();
				//country check
				try {
					CC = WB.findElement(By.xpath("//td[text()='Countries']/following-sibling::td")).getText();
				}
				catch(Exception cce) {
					CC ="No Data Found";
				}
				List<WebElement> No_child_check_path=new ArrayList<WebElement>();
				No_child_check_path=WB.findElements(By.xpath("//td[text()='Level " + (Level+1) + " (Child)']"));
				if(!CC.contains("No Data Found")) {
					if(No_child_check_path.size()==0) {
						System.out.println(WB.getCurrentUrl());
						if(Level==2) {
							E_L2=Ranges;
						} else if(Level==3) {
							E_L3=Ranges;
						} else if(Level==4) {
							E_L4=Ranges;
						} else if(Level==5) {
							E_L5=Ranges;
						} else if(Level==6) {
							E_L6=Ranges;
						}								
			
						processTable(WB,E_L2,E_L3,E_L4,E_L5,E_L6,CC);
						WB.findElement(By.xpath(Parent_Path)).click();
					}
					else {
						LvlRngeLst.add(Ranges);
						WB.findElement(By.xpath(Parent_Path)).click();
					}
					
				}
				else
				{
					WB.findElement(By.xpath(Parent_Path)).click();
				}
			}
		}

		return LvlRngeLst;
	}

	//Setting the driver properties as the Browser selected by the user
	static void driverProperties(String msg, String driverPath_arg, String driverProperty_arg) {
		System.out.println("Setting the Driver Properties for execution");
		System.out.format("Running the code with %s Browser \n", msg);
		System.setProperty(driverProperty_arg, driverPath_arg);
	}
	
	//Launching  the User selected Browser
	static WebDriver LaunchBrowser(String Browser,String DriverExePath) {
		WebDriver driverClass = null;
		if (Browser.contains("Chrome")) {			
			driverProperties("Chrome", DriverExePath + "chromedriver.exe", "webdriver.chrome.driver");
			driverClass = new ChromeDriver();
		} else if (Browser.contains("Firefox")) {
			driverProperties("Firefox", DriverExePath + "geckodriver.exe", "webdriver.gecko.driver");
			driverClass = new FirefoxDriver();
		} else if (Browser.contains("Edge")) {
			driverProperties("Edge", DriverExePath + "msedgedriver.exe", "webdriver.edge.driver");
			driverClass = new EdgeDriver();
		}
		return driverClass;
	}
	
	static void processTable(WebDriver WD,String L2,String L3,String L4,String L5,String L6,String Country) throws IOException {
		String path=".\\RequiredFiles\\NorthAmerica_Mountain_ranges.xlsx";
		FileInputStream fs = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		String table_xpath="//table[@class='gray'][3]//tr";
		int row_size = WD.findElements(By.xpath(table_xpath)).size();
		System.out.println("Table row size - " + row_size);
		
		for(int cntr=3;cntr<row_size;cntr++) {
			Row row1 = sheet.createRow(Globals.ExcelRowCount);
			int colno=6;
			for(int i=1;i<5;i++) {
				String col_xpath="//table[@class='gray'][3]//tr["+ cntr + "]//td[" + i + "]";
				String Celltext = WD.findElement(By.xpath(col_xpath)).getText();
				Cell c1 = row1.createCell(colno);
				c1.setCellValue(Celltext);				
				colno++;
				if(i==2) {
					String HyperLnk_xpath = "//table[@class='gray'][3]//tr["+ cntr + "]//td[" + i + "]/a";
					String Lnk_url = WD.findElement(By.xpath(HyperLnk_xpath)).getAttribute("href");
					System.out.println(Lnk_url);
					Cell c12 = row1.createCell(12);
					c12.setCellValue(Lnk_url);
				}
				Cell c2 = row1.createCell(1);
				c2.setCellValue(L2);
				Cell c3 = row1.createCell(2);
				c3.setCellValue(L3);
				Cell c4 = row1.createCell(3);
				c4.setCellValue(L4);
				Cell c5 = row1.createCell(4);
				c5.setCellValue(L5);
				Cell c6 = row1.createCell(5);
				c6.setCellValue(L6);
				Cell c7=row1.createCell(13);
				c7.setCellValue(Country);
				System.out.println("Cellvalue " + i + ": " + Celltext);				
			}
			Globals.ExcelRowCount++;
		}
		fs.close();
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();
	}


}
