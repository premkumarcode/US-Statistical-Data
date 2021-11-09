
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.JavascriptExecutor;

//Working Perfectly - No Changes to be made

public class US_StatisticalData_Drugoverdose {

	public static void main(String[] args) throws InterruptedException, IOException, NullPointerException {
		String Path = ".\\BrowserDriver\\";
		WebDriver driverClass = null;

		System.out.println("Enter the Browser choice ['Chrome' / 'Firefox' / 'Edge'] \n");
		// Scanner scan = new Scanner(System.in);
		// String driverToRun=scan.nextLine();
		// scan.close();
		String driverToRun = "Chrome";
		String path = ".\\RequiredFiles\\Drug_overdose_deaths_US.xlsx";

		// open the excel file to read
		FileInputStream fs = new FileInputStream(path);
		// Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);

		if (driverToRun.contains("Chrome")) {
			System.out.println(Path);
			driverProperties("Chrome", Path + "chromedriver.exe", "webdriver.chrome.driver");
			driverClass = new ChromeDriver();
		} else if (driverToRun.contains("Firefox")) {
			driverProperties("Firefox", Path + "geckodriver.exe", "webdriver.gecko.driver");
			driverClass = new FirefoxDriver();
		} else if (driverToRun.contains("Edge")) {
			driverProperties("Edge", Path + "msedgedriver.exe", "webdriver.edge.driver");
			driverClass = new EdgeDriver();
		}

		JavascriptExecutor js = (JavascriptExecutor) driverClass;
		int RowCount = 3;
		WebDriverWait wait = new WebDriverWait(driverClass, 10);
		String url_Navig = "https://opioidmisusetool.norc.org/";
		driverClass.get(url_Navig);
		driverClass.manage().window().maximize();
		Thread.sleep(2000);
		// Click on the County List
		driverClass.findElement(By.xpath("//div[@id='countyListButton']")).click();

		for (int cntr = 1; cntr <= 51; cntr++) {
			if (cntr != 8) {
				String xpathVar1 = "//div[@class='button-container'][5]//div[@class='miniMenu'][" + cntr + "]/div";
				System.out.println(xpathVar1);

				WebElement element = driverClass.findElement(By.xpath(xpathVar1));
//					//js.executeScript("arguments[0].scrollIntoView();", element);
//					if((element.isDisplayed()==false) && (element.isEnabled()==false) ) {
//						System.out.println("I am inside county list click");
//						Thread.sleep(2000);
//						js.executeScript("arguments[0].scrollIntoView();", element);
//					}
				Thread.sleep(2000);
				Boolean status_state = false;
				int state_click_counter = 0;
				// click on the county name to open the county details window and activate the
				// popup window
				while ((status_state == false) || (state_click_counter < 20)) {
					System.out.println("Inside the loop - " + status_state + " ----" + state_click_counter);
					try {
						element.click();
						status_state = true;
						state_click_counter = 21;
					} catch (Exception e) {
						js.executeScript("arguments[0].scrollIntoView();", element);
						state_click_counter++;

					}
				}

				String col1_State_Name = element.getText();
				String xpathVar2 = "//div[@class='button-container'][5]//div[@class='miniMenu'][" + cntr
						+ "]/button[@class='menuButton menuDropdownElement countyDropDownElement']";
				List<WebElement> countylist = driverClass.findElements(By.xpath(xpathVar2));

				for (int cntr1 = 0; cntr1 < countylist.size(); cntr1++) {
					System.out.println("I am inside for county list - " + countylist.size());
					System.out.println(col1_State_Name);
					String col2_County_Name = countylist.get(cntr1).getText();
					System.out.println(col2_County_Name);
					Thread.sleep(2000);
					Boolean status_county = false;
					int county_click_counter = 0;
					// click on the county name to open the county details window and activate the
					// popup window
					while (status_county == false || county_click_counter < 20) {
						try {
							countylist.get(cntr1).click();
							status_county = true;
							county_click_counter = 21;
						} catch (Exception e) {
							js.executeScript("arguments[0].scrollIntoView();", countylist.get(cntr1));
							county_click_counter++;

						}
					}

					driverClass.switchTo().activeElement();

					// Creating row in the excel and writing the value to excel
					Row row1 = sheet.createRow(RowCount);
					Cell c1 = row1.createCell(0);
					c1.setCellValue(col1_State_Name);
					Cell c2 = row1.createCell(1);
					c2.setCellValue(col2_County_Name);

					WebElement closebtn1 = driverClass
							.findElement(By.xpath("//div[@id='countyDetail']/div[@class='closediv']/span/span"));

					Boolean temp = closebtn1.isDisplayed();
					System.out.println(temp);
					if (temp == false) {
						System.out.println("I am here inside exception");
						Boolean closebtn_status = false;
						while (closebtn_status == false) {
							driverClass.get(url_Navig);
							Thread.sleep(2000);
							// Click on the County List
							driverClass.findElement(By.xpath("//div[@id='countyListButton']")).click();
							Thread.sleep(2000);
							status_state = false;
							state_click_counter = 0;
							// click on the county name to open the county details window and activate the
							// popup window
							while ((status_state == false) || (state_click_counter < 20)) {
								System.out.println("Inside the loop - " + status_state + " ----" + state_click_counter);
								try {
									element = driverClass.findElement(By.xpath(xpathVar1));
									element.click();
									status_state = true;
									state_click_counter = 21;
								} catch (Exception e1) {
									js.executeScript("arguments[0].scrollIntoView();", element);
									state_click_counter++;

								}
							}
							Thread.sleep(2000);
							status_county = false;
							county_click_counter = 0;
							// click on the county name to open the county details window and activate the
							// popup window
							while (status_county == false || county_click_counter < 20) {
								try {
									countylist = driverClass.findElements(By.xpath(xpathVar2));
									countylist.get(cntr1).click();
									status_county = true;
									county_click_counter = 21;
								} catch (Exception e2) {
									js.executeScript("arguments[0].scrollIntoView();", countylist.get(cntr1));
									county_click_counter++;

								}
							}

							driverClass.switchTo().activeElement();
							closebtn_status = true;
							break;

						}
					}
					String col3_dp_Drate = driverClass.findElement(By.xpath("//td[@id='dp_Drate']")).getText();
					Cell c3 = row1.createCell(2);
					c3.setCellValue(col3_dp_Drate);

					String col4_ap_Drate = driverClass.findElement(By.xpath("//td[@id='ap_Drate']")).getText();
					Cell c4 = row1.createCell(3);
					c4.setCellValue(col4_ap_Drate);

					String col5_dp_Deaths = driverClass.findElement(By.xpath("//td[@id='dp_Deaths']")).getText();
					Cell c5 = row1.createCell(4);
					c5.setCellValue(col5_dp_Deaths);

					String col6_dp_TotalPop = driverClass.findElement(By.xpath("//td[@id='dp_TotalPop']")).getText();
					Cell c6 = row1.createCell(5);
					c6.setCellValue(col6_dp_TotalPop);

					String col7_dp_Urbanicity = driverClass.findElement(By.xpath("//td[@id='dp_Urbanicity']"))
							.getText();
					Cell c7 = row1.createCell(6);
					c7.setCellValue(col7_dp_Urbanicity);

					String col8_dp_Pct_Wht = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Wht']")).getText();
					Cell c8 = row1.createCell(7);
					c8.setCellValue(col8_dp_Pct_Wht);

					String col9_ap_Pct_Wht = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Wht']")).getText();
					Cell c9 = row1.createCell(8);
					c9.setCellValue(col9_ap_Pct_Wht);

					String col9a_us_Pct_Wht = driverClass.findElement(By.xpath("//td[@id='us_Pct_Wht']")).getText();
					Cell c9a = row1.createCell(9);
					c9a.setCellValue(col9a_us_Pct_Wht);

					String col10_dp_Pct_Blk = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Blk']")).getText();
					Cell c10 = row1.createCell(10);
					c10.setCellValue(col10_dp_Pct_Blk);

					String col11_ap_Pct_Blk = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Blk']")).getText();
					Cell c11 = row1.createCell(11);
					c11.setCellValue(col11_ap_Pct_Blk);

					String col12_us_Pct_Blk = driverClass.findElement(By.xpath("//td[@id='us_Pct_Blk']")).getText();
					Cell c12 = row1.createCell(12);
					c12.setCellValue(col12_us_Pct_Blk);

					String col13_dp_Pct_Hisp = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Hisp']")).getText();
					Cell c13 = row1.createCell(13);
					c13.setCellValue(col13_dp_Pct_Hisp);

					String col14_ap_Pct_Hisp = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Hisp']")).getText();
					Cell c14 = row1.createCell(14);
					c14.setCellValue(col14_ap_Pct_Hisp);

					String col15_us_Pct_Hisp = driverClass.findElement(By.xpath("//td[@id='us_Pct_Hisp']")).getText();
					Cell c15 = row1.createCell(15);
					c15.setCellValue(col15_us_Pct_Hisp);

					String col16_dp_Pct_Asn = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Asn']")).getText();
					Cell c16 = row1.createCell(16);
					c16.setCellValue(col16_dp_Pct_Asn);

					String col17_ap_Pct_Asn = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Asn']")).getText();
					Cell c17 = row1.createCell(17);
					c17.setCellValue(col17_ap_Pct_Asn);

					String col18_us_Pct_Asn = driverClass.findElement(By.xpath("//td[@id='us_Pct_Asn']")).getText();
					Cell c18 = row1.createCell(18);
					c18.setCellValue(col18_us_Pct_Asn);

					String col19_dp_Pct_Pac = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Pac']")).getText();
					Cell c19 = row1.createCell(19);
					c19.setCellValue(col19_dp_Pct_Pac);

					String col20_ap_Pct_Pac = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Pac']")).getText();
					Cell c20 = row1.createCell(20);
					c20.setCellValue(col20_ap_Pct_Pac);

					String col21_us_Pct_Pac = driverClass.findElement(By.xpath("//td[@id='us_Pct_Pac']")).getText();
					Cell c21 = row1.createCell(21);
					c21.setCellValue(col21_us_Pct_Pac);

					String col22_dp_Pct_Nat = driverClass.findElement(By.xpath("//td[@id='dp_Pct_Nat']")).getText();
					Cell c22 = row1.createCell(22);
					c22.setCellValue(col22_dp_Pct_Nat);

					String col23_ap_Pct_Nat = driverClass.findElement(By.xpath("//td[@id='ap_Pct_Nat']")).getText();
					Cell c23 = row1.createCell(23);
					c23.setCellValue(col23_ap_Pct_Nat);

					String col24_us_Pct_Nat = driverClass.findElement(By.xpath("//td[@id='us_Pct_Nat']")).getText();
					Cell c24 = row1.createCell(24);
					c24.setCellValue(col24_us_Pct_Nat);

					String col25_dp_Under15 = driverClass.findElement(By.xpath("//td[@id='dp_Under15']")).getText();
					Cell c25 = row1.createCell(25);
					c25.setCellValue(col25_dp_Under15);

					String col26_ap_Under15 = driverClass.findElement(By.xpath("//td[@id='ap_Under15']")).getText();
					Cell c26 = row1.createCell(26);
					c26.setCellValue(col26_ap_Under15);

					String col27_us_Under15 = driverClass.findElement(By.xpath("//td[@id='us_Under15']")).getText();
					Cell c27 = row1.createCell(27);
					c27.setCellValue(col27_us_Under15);

					String col28_dp_F15to65 = driverClass.findElement(By.xpath("//td[@id='dp_F15to65']")).getText();
					Cell c28 = row1.createCell(28);
					c28.setCellValue(col28_dp_F15to65);

					String col29_ap_F15to65 = driverClass.findElement(By.xpath("//td[@id='ap_F15to65']")).getText();
					Cell c29 = row1.createCell(29);
					c29.setCellValue(col29_ap_F15to65);

					String col30_us_F15to65 = driverClass.findElement(By.xpath("//td[@id='us_F15to65']")).getText();
					Cell c30 = row1.createCell(30);
					c30.setCellValue(col30_us_F15to65);

					String col31_dp_Over65 = driverClass.findElement(By.xpath("//td[@id='dp_Over65']")).getText();
					Cell c31 = row1.createCell(31);
					c31.setCellValue(col31_dp_Over65);

					String col32_ap_Over65 = driverClass.findElement(By.xpath("//td[@id='ap_Over65']")).getText();
					Cell c32 = row1.createCell(32);
					c32.setCellValue(col32_ap_Over65);

					String col33_us_Over65 = driverClass.findElement(By.xpath("//td[@id='us_Over65']")).getText();
					Cell c33 = row1.createCell(33);
					c33.setCellValue(col33_us_Over65);

					String col34_dp_Comp_HS = driverClass.findElement(By.xpath("//td[@id='dp_Comp_HS']")).getText();
					Cell c34 = row1.createCell(34);
					c34.setCellValue(col34_dp_Comp_HS);

					String col35_ap_Comp_HS = driverClass.findElement(By.xpath("//td[@id='ap_Comp_HS']")).getText();
					Cell c35 = row1.createCell(35);
					c35.setCellValue(col35_ap_Comp_HS);

					String col36_us_Comp_HS = driverClass.findElement(By.xpath("//td[@id='us_Comp_HS']")).getText();
					Cell c36 = row1.createCell(36);
					c36.setCellValue(col36_us_Comp_HS);

					String col37_dp_Comp_CO = driverClass.findElement(By.xpath("//td[@id='dp_Comp_CO']")).getText();
					Cell c37 = row1.createCell(37);
					c37.setCellValue(col37_dp_Comp_CO);

					String col38_ap_Comp_CO = driverClass.findElement(By.xpath("//td[@id='ap_Comp_CO']")).getText();
					Cell c38 = row1.createCell(38);
					c38.setCellValue(col38_ap_Comp_CO);

					String col39_us_Comp_CO = driverClass.findElement(By.xpath("//td[@id='us_Comp_CO']")).getText();
					Cell c39 = row1.createCell(39);
					c39.setCellValue(col39_us_Comp_CO);

					String col40_dp_Disable = driverClass.findElement(By.xpath("//td[@id='dp_Disable']")).getText();
					Cell c40 = row1.createCell(40);
					c40.setCellValue(col40_dp_Disable);

					String col41_ap_Disable = driverClass.findElement(By.xpath("//td[@id='ap_Disable']")).getText();
					Cell c41 = row1.createCell(41);
					c41.setCellValue(col41_ap_Disable);

					String col42_us_Disable = driverClass.findElement(By.xpath("//td[@id='us_Disable']")).getText();
					Cell c42 = row1.createCell(42);
					c42.setCellValue(col42_us_Disable);

					String col43_dp_three_or_m = driverClass.findElement(By.xpath("//td[@id='dp_three_or_m']"))
							.getText();
					Cell c43 = row1.createCell(43);
					c43.setCellValue(col43_dp_three_or_m);

					String col44_ap_three_or_m = driverClass.findElement(By.xpath("//td[@id='ap_three_or_m']"))
							.getText();
					Cell c44 = row1.createCell(44);
					c44.setCellValue(col44_ap_three_or_m);

					String col45_us_three_or_m = driverClass.findElement(By.xpath("//td[@id='us_three_or_m']"))
							.getText();
					Cell c45 = row1.createCell(45);
					c45.setCellValue(col45_us_three_or_m);

					String col46_dp_MHHINC = driverClass.findElement(By.xpath("//td[@id='dp_MHHINC']")).getText();
					Cell c46 = row1.createCell(46);
					c46.setCellValue(col46_dp_MHHINC);

					String col47_ap_MHHINC = driverClass.findElement(By.xpath("//td[@id='ap_MHHINC']")).getText();
					Cell c47 = row1.createCell(47);
					c47.setCellValue(col47_ap_MHHINC);

					String col48_us_MHHINC = driverClass.findElement(By.xpath("//td[@id='us_MHHINC']")).getText();
					Cell c48 = row1.createCell(48);
					c48.setCellValue(col48_us_MHHINC);

					String col49_dp_Poverty = driverClass.findElement(By.xpath("//td[@id='dp_Poverty']")).getText();
					Cell c49 = row1.createCell(49);
					c49.setCellValue(col49_dp_Poverty);

					String col50_ap_Poverty = driverClass.findElement(By.xpath("//td[@id='ap_Poverty']")).getText();
					Cell c50 = row1.createCell(50);
					c50.setCellValue(col50_ap_Poverty);

					String col51_us_Poverty = driverClass.findElement(By.xpath("//td[@id='us_Poverty']")).getText();
					Cell c51 = row1.createCell(51);
					c51.setCellValue(col51_us_Poverty);

					String col52_dp_UnEmp_ = driverClass.findElement(By.xpath("//td[@id='dp_UnEmp_']")).getText();
					Cell c52 = row1.createCell(52);
					c52.setCellValue(col52_dp_UnEmp_);

					String col53_ap_UnEmp_ = driverClass.findElement(By.xpath("//td[@id='ap_UnEmp_']")).getText();
					Cell c53 = row1.createCell(53);
					c53.setCellValue(col53_ap_UnEmp_);

					String col54_us_UnEmp_ = driverClass.findElement(By.xpath("//td[@id='us_UnEmp_']")).getText();
					Cell c54 = row1.createCell(54);
					c54.setCellValue(col54_us_UnEmp_);

					WebElement element55 = driverClass.findElement(By.xpath("//td[@id='dp_Constr_']"));
					js.executeScript("arguments[0].scrollIntoView();", element55);
					String col55_dp_Constr = driverClass.findElement(By.xpath("//td[@id='dp_Constr_']")).getText();
					Cell c55 = row1.createCell(55);
					c55.setCellValue(col55_dp_Constr);

					String col56_ap_Constr = driverClass.findElement(By.xpath("//td[@id='ap_Constr_']")).getText();
					Cell c56 = row1.createCell(56);
					c56.setCellValue(col56_ap_Constr);

					String col57_us_Constr = driverClass.findElement(By.xpath("//td[@id='us_Constr_']")).getText();
					Cell c57 = row1.createCell(57);
					c57.setCellValue(col57_us_Constr);

					String col58_dp_Mining = driverClass.findElement(By.xpath("//td[@id='dp_Mining']")).getText();
					Cell c58 = row1.createCell(58);
					c58.setCellValue(col58_dp_Mining);

					String col59_ap_Mining = driverClass.findElement(By.xpath("//td[@id='ap_Mining']")).getText();
					Cell c59 = row1.createCell(59);
					c59.setCellValue(col59_ap_Mining);

					String col60_us_Mining = driverClass.findElement(By.xpath("//td[@id='us_Mining']")).getText();
					Cell c60 = row1.createCell(60);
					c60.setCellValue(col60_us_Mining);

					String col61_dp_Manu_ = driverClass.findElement(By.xpath("//td[@id='dp_Manu_']")).getText();
					Cell c61 = row1.createCell(61);
					c61.setCellValue(col61_dp_Manu_);

					String col62_ap_Manu_ = driverClass.findElement(By.xpath("//td[@id='ap_Manu_']")).getText();
					Cell c62 = row1.createCell(62);
					c62.setCellValue(col62_ap_Manu_);

					String col63_us_Manu_ = driverClass.findElement(By.xpath("//td[@id='us_Manu_']")).getText();
					Cell c63 = row1.createCell(63);
					c63.setCellValue(col63_us_Manu_);

					String col64_dp_Trade = driverClass.findElement(By.xpath("//td[@id='dp_Trade']")).getText();
					Cell c64 = row1.createCell(64);
					c64.setCellValue(col64_dp_Trade);

					String col65_ap_Trade = driverClass.findElement(By.xpath("//td[@id='ap_Trade']")).getText();
					Cell c65 = row1.createCell(65);
					c65.setCellValue(col65_ap_Trade);

					String col66_us_Trade = driverClass.findElement(By.xpath("//td[@id='us_Trade']")).getText();
					Cell c66 = row1.createCell(66);
					c66.setCellValue(col66_us_Trade);

					FileOutputStream fos = new FileOutputStream(path);
					workbook.write(fos);
					fos.close();
					// Close the county detail popup window
					WebElement closebtn = driverClass
							.findElement(By.xpath("//div[@id='countyDetail']/div[@class='closediv']/span/span"));
					if (closebtn.isDisplayed() == false) {
						js.executeScript("arguments[0].scrollIntoView();", closebtn);
					}
					closebtn.click();
					Thread.sleep(4000);
					RowCount++;

				}
				js.executeScript("arguments[0].scrollIntoView();", element);
				Thread.sleep(4000);
				element.click();
				Thread.sleep(2000);
			}

		}
		Thread.sleep(1000);
		driverClass.close();
		fs.close();
	}

	static void driverProperties(String msg, String driverPath_arg, String driverProperty_arg) {
		System.out.println("Setting the Driver Properties for execution");
		System.out.format("Running the code with %s Browser \n", msg);
		System.setProperty(driverProperty_arg, driverPath_arg);
	}
}
