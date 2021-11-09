import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class US_StatisticalData_USASpending {

	public static void main(String[] args) throws InterruptedException, IOException {
		//Variable Declaration
		String BrowserDriverPath = ".\\BrowserDriver\\";
		String Browser_to_launch="Chrome";
		String WebsiteLaunched="https://www.usaspending.gov/state";
		
		String path=".\\RequiredFiles\\USASpending.xlsx";
		FileInputStream fs = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		
		WebDriver BrwsrLaunch=LaunchBrowser(Browser_to_launch,BrowserDriverPath);				
		BrwsrLaunch.get(WebsiteLaunched);
		BrwsrLaunch.manage().window().maximize();
		String parent=BrwsrLaunch.getWindowHandle();
		
		Thread.sleep(3000);
		List<WebElement>statelist=BrwsrLaunch.findElements(By.xpath("//tbody[contains(@class,\"state-list__body\")]//a"));
		String selectLinkOpeninNewTab = Keys.chord(Keys.CONTROL,Keys.RETURN);
		int excelrowcount=1;
		for(WebElement we : statelist) {
			we.sendKeys(selectLinkOpeninNewTab);
			Set<String>s=BrwsrLaunch.getWindowHandles();
			Iterator<String> I1= s.iterator();
			I1.next();
			BrwsrLaunch.switchTo().window(I1.next());
			Thread.sleep(5000);
			//get the state name
			String state_name=BrwsrLaunch.findElement(By.xpath("//h2[@class='state-overview__title']")).getText();
			if(state_name.contains("Texas")) {
				Thread.sleep(5000);
			}
			System.out.println(state_name);
			List<WebElement>spend_data=BrwsrLaunch.findElements(By.xpath("//*[@class='bar-data']//*[contains(text(),'Spending')]"));
			for(WebElement we2 : spend_data) {
				Row row1 = sheet.createRow(excelrowcount);
				String[] spenddata=we2.getText().split(":");
				String amount = spenddata[1].trim();
				String year = spenddata[0].replace("Spending in ", "").trim();
				
				Cell c1 = row1.createCell(0);
				c1.setCellValue(state_name);
				
				Cell c2 = row1.createCell(1);
				c2.setCellValue(year);
				
				Cell c3 = row1.createCell(2);
				c3.setCellValue(amount);
				
				excelrowcount=excelrowcount+1;
				
				FileOutputStream fos = new FileOutputStream(path);
				workbook.write(fos);
				fos.close();
				
			}
			

			BrwsrLaunch.close();
			BrwsrLaunch.switchTo().window(parent);
		
			
		}
		fs.close();

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
	
	static void driverProperties(String msg,String driverPath_arg,String driverProperty_arg) {
		  System.out.println("Setting the Driver Properties for execution");
		  System.out.format("Running the code with %s Browser \n",msg);
		  System.setProperty(driverProperty_arg,driverPath_arg);		  
		}
}
