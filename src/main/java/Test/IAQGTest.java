package Test;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import javax.swing.JOptionPane;
import javax.swing.UIManager;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class IAQGTest {

//Variable declaration
	String strUrl = null,strNotepadFilePath = null,strPropertyFilePath = "C:\\Users\\DongareS\\Desktop\\IAQG\\Data\\Input.properties";
	Properties propClass;
	int counter =1;

//Functions
	
	//Read property File
		public Properties readPropertyFile(String propertyFilePath)
		{
			Properties p = new Properties();
			String strFileName= propertyFilePath;
			BufferedReader b = null;
			
			try {
				b = new BufferedReader(new FileReader(strFileName));
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			try {
				p.load(b);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			try {
				b.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}	
			return p;
		}

	public void createTextFile()
	{
		File myObj = new File(strNotepadFilePath);
		if (myObj.exists()) {
			myObj.delete();
			try {
				myObj.createNewFile();
			} catch (IOException e) {
			// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.println("Existing file deleted and created new one: " + myObj.getName());
		} else {
		try {
			myObj.createNewFile();
			System.out.println("New file created.");
		} catch (IOException e) {
		// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		}
	}

	//Write data to Notepad
	public void writeToText(String data)
	{
		try (FileWriter f = new FileWriter(strNotepadFilePath, true); 
			BufferedWriter b = new BufferedWriter(f); 
			PrintWriter p = new PrintWriter(b);) 
			{
			p.println(counter+". "+data);
			counter=counter+1;
		}
		catch (IOException i)
		{ i.printStackTrace(); 
		}
	}

	//trim string from left side
	public static String ltrim(String s) {
		int i = 0;
		while (i < s.length() && Character.isWhitespace(s.charAt(i))) {
			i++;
		}
		return s.substring(i);
	}

	//Fetch company URL from google
	public void getCompanyURL(WebDriver driver,String strCompanyName,String strParentWindow,String strCountry)
	{
		((JavascriptExecutor)driver).executeScript("window.open()");
		String urlToOpen = "https://www.google.com";
		String strChild;
		Set<String> childs = driver.getWindowHandles();
		Iterator ic = childs.iterator();
		while(ic.hasNext())
		{								            	   
			strChild = (String)ic.next();
			if(!strChild.equalsIgnoreCase(strParentWindow))
			{
				driver.switchTo().window(strChild);
				driver.get(urlToOpen);
				WebElement wGoogle = driver.findElement(By.xpath("//*[@id=\"tsf\"]//input[@class=\"gLFyf gsfi\"]"));					

				wGoogle.sendKeys(strCompanyName +"," +strCountry);	
				wGoogle.sendKeys(Keys.ENTER);														
				try {
					WebElement childUrl = driver.findElement(By.xpath("//div[text()='Website']/parent::a"));	
					strUrl =childUrl.getAttribute("href");
				}	catch(Exception e) {
					if ( !(e instanceof NoSuchElementException) ) {
						System.out.println(e);
						
						wGoogle.sendKeys(strCompanyName);	
						wGoogle.sendKeys(Keys.ENTER);	
						try {
							WebElement childUrl = driver.findElement(By.xpath("//div[text()='Website']/parent::a"));	
							strUrl =childUrl.getAttribute("href");
						}	catch(Exception e1) {
							if ( !(e1 instanceof NoSuchElementException) ) {
								System.out.println(e1);
							}
						}
					}
				}
			

			}
		}
	}

	//Get Integer property value	
	public Integer getInteger(String key)
	{
	   Integer value = null;
	   String string = propClass.getProperty(key);
	   if (string != null)
	      value = new Integer(string);
	   return value;
	}
	
//main	
	public static void main(String[] args) {

			IAQGTest objTestIAQG = new IAQGTest();	
			objTestIAQG.propClass = objTestIAQG.readPropertyFile(objTestIAQG.strPropertyFilePath);
			Properties prop = objTestIAQG.propClass;
			
			objTestIAQG.strNotepadFilePath = prop.getProperty("strNotepadFilePath");
			System.out.println("Notepad file path : "+objTestIAQG.strNotepadFilePath);
			objTestIAQG.createTextFile();
			int intFlag,srNo;
			srNo = objTestIAQG.getInteger("srNo");
			System.setProperty("webdriver.chrome.driver",prop.getProperty("strBrowserDriver"));
			WebDriver driver;
		
			//Excel code
			String fileName = prop.getProperty("strExcelFilePath");;		
			File f = new File(fileName);
		
//					if(f.exists()) {
//						f.delete();
//						System.out.println("Deleted");
//						try {
//							f.createNewFile();
//							System.out.println("Created");
//						} catch (IOException e) {
//							// TODO Auto-generated catch block
//							e.printStackTrace();
//						}
//					}else
//					{
//						try {
//							f.createNewFile();
//							System.out.println("Created");
//						} catch (IOException e) {
//							// TODO Auto-generated catch block
//							e.printStackTrace();
//						}
//					
//					}
		
			FileInputStream fis=null;
			int intRecord =0;
						
			try {
				fis = new FileInputStream(f);
			} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			XSSFWorkbook w = null;
			try {
				w = new XSSFWorkbook(fis);
			} catch (IOException e2) {
			// TODO Auto-generated catch block
				e2.printStackTrace();
			}
			XSSFSheet s =w.getSheet("Sheet1");
			Row r1;
			CellStyle cs = w.createCellStyle();
			cs.setWrapText(true);
			
			r1 =s.getRow(0);
			if( r1== null) {
				r1 = s.createRow(intRecord);				     
			}
			r1.createCell(0).setCellValue("Sr.No.");
			r1.createCell(1).setCellValue("Comapny Name");
			r1.createCell(2).setCellValue("Scope/Area of Business/ Services");
			r1.createCell(3).setCellValue("Contact Name");
			r1.createCell(4).setCellValue("Proffesional email");
			r1.createCell(5).setCellValue("Contact Extracted from");
			r1.createCell(6).setCellValue("Phone");
			r1.createCell(7).setCellValue("Office Location");
			r1.createCell(8).setCellValue("Website Name");
			r1.createCell(9).setCellValue("Customers Served");
			r1.createCell(10).setCellValue("About/Other information");
			r1.createCell(11).setCellValue("Co size");
			r1.createCell(12).setCellValue("Revenue");
			r1.createCell(13).setCellValue("Email Status");
			r1.createCell(14).setCellValue("Remark");
		
		
			//login code
			driver = new ChromeDriver();
			Actions action = new Actions(driver);
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			driver.get("https://www.iaqg.org/oasis/login");		
			driver.findElement(By.xpath("//*[@id=\"frm-login-1-input\"]")).sendKeys(prop.getProperty("strUsername"));
			driver.findElement(By.xpath("//*[@id=\"frm-login-2-input\"]")).sendKeys(prop.getProperty("strPassword"));
			driver.findElement(By.xpath("//*[@id=\"frm-login-3\"]/div[2]/a/div/div/span")).click();
		
		
			//navigation to company page
			driver.findElement(By.xpath("//*[@id=\"hdr-main-data\"]/a")).click();
			driver.findElement(By.xpath("//*[@id=\"hdr-main-data\"]/div/div[2]/ul/li[1]/a")).click();
			WebElement d1 = driver.findElement(By.xpath("//*[@id=\"sector\"]"));
			Select s1 = new Select(d1);		
			s1.selectByVisibleText(prop.getProperty("strSector"));		
			WebElement d2 = driver.findElement(By.xpath("//*[@id=\"country\"]"));
			Select s2 = new Select(d2);	
			String strCountryName = prop.getProperty("strCountry");
			s2.selectByVisibleText(strCountryName);
			driver.findElement(By.xpath("//*[@id=\"csdsearch_0\"]")).click();	
		
			//selection of specific page
			for (int k=objTestIAQG.getInteger("startPageNo");k<=objTestIAQG.getInteger("endPageNo");k++) {
		
				WebElement p = driver.findElement(By.xpath("//*[@id=\"csd--suppliers-list\"]//div[@class='frm-results-nav-pages']//a[text()="+k+"]"));	
				p.click();
				try {
					Thread.sleep(1500);
				} catch (InterruptedException e2) {
				// TODO Auto-generated catch block
					e2.printStackTrace();
				}
				List<WebElement> wCompany = driver.findElements(By.xpath("//*[@id=\"csd--suppliers-list\"]/div[3]/table/tbody/tr"));
				System.out.println("Company size "+wCompany.size());
				JavascriptExecutor js ;
				intFlag = 0 ;
				WebElement wCompanyName ;
			
				driver.manage().window().maximize();
				
				//Iteration over company to collect data
				for (int i=objTestIAQG.getInteger("startComapanyNo");i<=objTestIAQG.getInteger("endComapanyNo");i++)
				{
					try 
					{
						intFlag=0;
						objTestIAQG.strUrl = "";
						wCompanyName = driver.findElement(By.xpath("//*[@id=\"csd--suppliers-list\"]/div[3]/table/tbody/tr["+i+"]/td[1]/div[1]/div[3]/span/a/span/strong"));
						((JavascriptExecutor)driver).executeScript("arguments[0].scrollIntoView(false);", wCompanyName);
						Thread.sleep(800); 			
						wCompanyName.click();
					
						try {
							WebElement btnClose = driver.findElement(By.xpath("//*[@id=\"yes-btn\"]//span[text()=\"Yes\"]"));
							System.out.println("Warning message displayed");
							btnClose.isDisplayed();
							btnClose.click();
						}
						catch(Exception e)
						{
							System.out.println("Company profile opened successfully without warning message");
						}			
					
						try {
							Thread.sleep(800);
						} catch (InterruptedException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					
						js =(JavascriptExecutor)driver;
						js.executeScript("window.scrollTo(500, 500)");
						try {
							Thread.sleep(400);
						} catch (InterruptedException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					
						WebElement wScope =  driver.findElement(By.xpath("//*[@class=\"frmi frmi-w-full frmi-r-float\"]//div[@class=\"frmi-v-o\"]/ul/li"));
						String strScope = wScope.getText().replace("\n", ". ").replace("\r", ". ");	
						strScope =strScope.replaceAll( 
						"[^a-zA-Z0-9 ]", "")+"."; 
						strScope =ltrim(strScope);
					
						js =(JavascriptExecutor)driver;
						js.executeScript("window.scrollTo(0, 0)");
						Thread.sleep(100);
						String strCompanyName = driver.findElement(By.xpath("//*[@id=\"watchlist-toggle\"]/strong")).getText();
						String strAddress =driver.findElement(By.xpath("//*[@id=\"main\"]//div[@class=\"frm-org frm-org-info\"]//div[@class=\"ct-address\"]")).getText();
						String strOIN=driver.findElement(By.xpath("//*[@id=\"main\"]//em[text()=\"OIN: \"]/span")).getText();				
					
						String strCriteria = "Manufacture~mechanical~Aerospace~Aeronautics~Automotive~Banking & Financial Services~Energy and Utilities~Gaming~Healthcare & Pharmaceutical~Industrial~Insurance~Naval~Public Sector~Rail Transportation~Retail & Logistics~defence~automotive~industry & manufacturing~medical devices~banking and financial services~insurance and retail~space & Defence~Telecoms & Media~manufacturing~business consulting~design~analysis~quality assurance~quality engineering~process automation~Testing~Automation Testing~Electronics~Electrical~Design~Development~Production";
						String[] arCriteria =strCriteria.split("~");
						for(int j=0;j<arCriteria.length;j++) {
							if (strScope.toLowerCase().contains(arCriteria[j].toLowerCase()))
							{
								intFlag=1;
								System.out.println(strCompanyName+" - "+strScope);
							}
					
						}
					
						if(intFlag==1){
							intRecord = intRecord+1;							
							r1 =s.getRow(intRecord);
							if( r1== null) {
								r1 = s.createRow(intRecord);				     
							}
						
							r1.createCell(0).setCellValue(srNo);
							r1.createCell(1).setCellValue(strCompanyName);
							r1.createCell(2).setCellValue(strScope);
							r1.createCell(3).setCellValue("");
							r1.createCell(4).setCellValue("");
							r1.createCell(5).setCellValue("");
							r1.createCell(6).setCellValue("");
							r1.createCell(7).setCellValue(strAddress);
						
							String strParentWindow = driver.getWindowHandle();
							//Checking if contactinfo link present
							try {
								driver.findElement(By.xpath("//*[@id=\"main\"]/div[4]/div[1]/div/div/div/div[1]/a")).isDisplayed();							
							
								if (driver.findElement(By.xpath("//*[@id=\"main\"]/div[4]/div[1]/div/div/div/div[1]/a")).isEnabled())
								{
									driver.findElement(By.xpath("//*[@id=\"main\"]/div[4]/div[1]/div/div/div/div[1]/a")).click();
									try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									String strContact = driver.findElement(By.xpath("//*[@id=\"overlay-content\"]/div[2]/div[2]/div/div[1]/span/div[2]/div/ul/li/span")).getText();
									String strPhone =driver.findElement(By.xpath("//*[@id=\"overlay-content\"]/div[2]/div[2]/div/div[2]/span/div[2]/div/ul/li/span")).getText();
									String strEmail =driver.findElement(By.xpath("//*[@id=\"overlay-content\"]/div[2]/div[2]/div/div[4]/span/div[2]/div/ul/li/span/a")).getText();
								
									r1.createCell(3).setCellValue(strContact);
									r1.createCell(4).setCellValue(strEmail);
									r1.createCell(5).setCellValue("IAQG");
									r1.createCell(6).setCellValue(strPhone);
									//r1.createCell(7).setCellStyle(cs);				
								
								
									driver.findElement(By.xpath("//*[@id=\"overlay-content\"]//a[text()=\"Web Site\"]")).click();
								
									Set<String> strChildWindows = driver.getWindowHandles();
									Iterator r = strChildWindows.iterator();
									String childWindow;	
								
									while(r.hasNext())
									{
										childWindow =(String) r.next();								
										if(!childWindow.equalsIgnoreCase(strParentWindow))				
										{
											driver.switchTo().window(childWindow);
											objTestIAQG.strUrl =driver.getCurrentUrl();
										
											if (objTestIAQG.strUrl.trim().contains("about:blank#blocked")) {
										
												driver.close();
												driver.switchTo().window(strParentWindow);				 
											
												objTestIAQG.getCompanyURL(driver, strCompanyName, strParentWindow,strCountryName);
												System.out.println("href value = "+objTestIAQG.strUrl);	
											
											}
										
											r1.createCell(8).setCellValue(objTestIAQG.strUrl);	
											driver.close();
										}
									
									}
									driver.switchTo().window(strParentWindow);
									Thread.sleep(400);
									driver.findElement(By.xpath("//*[@id=\"overlay\"]/div/div[2]/a")).click();
								
									try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								}
							}
							catch(Exception e) {
							if ( !(e instanceof NoSuchElementException) ) {
								System.out.println(e);
							}	
								objTestIAQG.getCompanyURL(driver, strCompanyName, strParentWindow,strCountryName);
							
								r1.createCell(8).setCellValue(objTestIAQG.strUrl);
								driver.close();
								driver.switchTo().window(strParentWindow);
								Thread.sleep(400);
							}
							srNo = srNo+1;
						}
						else
						{
							System.out.println("Sorry ,scope not matching"+strCompanyName +" - "+strScope);
							objTestIAQG.writeToText(strCompanyName +" - "+strScope);
						
						}
						driver.findElement(By.xpath("//*[@id=\"brd-return\"]/a")).click();
					}
					catch(Exception e)
					{
						System.out.println(e);
					//driver.close();
					}
					if (i==objTestIAQG.getInteger("endComapanyNo"))
					{
						System.out.println("Run Ended");
						JOptionPane.showMessageDialog(null,"Ended with records entered","Info",JOptionPane.INFORMATION_MESSAGE);
					}
				}
			
				FileOutputStream fos=null;
				try {
					fos = new FileOutputStream(f);
				} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				try {
					w.write(fos);
				} catch (IOException e) {
				// TODO Auto-generated catch block
					e.printStackTrace();
				}
				try {
					fos.close();
				} catch (IOException e) {
				// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			}	
	}		
}
