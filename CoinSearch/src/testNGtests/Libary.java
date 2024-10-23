
package testNGtests;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.lang.Boolean;

import javax.imageio.ImageIO;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.Select;


import org.testng.annotations.BeforeTest;

//import com.sun.tools.javac.util.Options;

import Operation.TakeScreenShot;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.read.biff.BiffException;
import java.io.FileInputStream;
import jxl.Sheet;


public class Libary
{	
	public static WritableWorkbook wwbCopy,wxCopy;
	public static WritableSheet wshTemp;
	public WebDriver driver;
	public String cururl,cururl1,url1,url3,Coin_Found_yr,Cuntry,yr_old,no,yr,s,url,currentYear,CCValue,comment,KM,PgNo1,C_VG,C_F,C_VF,C_XF,C_AU,C_UNC,C_VG_new,C_F_new,C_VF_new,C_XF_new,C_AU_new,C_UNC_new,Full_CCValue;
	public int index=0,i=0,rnum1,nrow,count=1,j=0,q=1,p,z=1,nop,blank,foo,foo1,Cctable;
	public Boolean isChecked,avail1,avail2,Coin_Yr_Avail,ImageCapture,avail=false,SameYr_Found,Coin_Found;
	public String links[] = new String[200];
	public String[] proofYears;

	
	public static void createExcelCountry ( String country1) throws IOException,RowsExceededException,WriteException,InterruptedException
	{
		//DateFormat dateFormat = new SimpleDateFormat("ddMMyyyyHHmmss");
		SimpleDateFormat dateFormat = new SimpleDateFormat("ddMMyyyyHHmmss");
		Date date = new Date ();
		String date1= dateFormat.format(date);
		wwbCopy = Workbook.createWorkbook(new File(System.getProperty("user.dir")+"\\Result\\"+country1+date1+".xls"));
		wshTemp = wwbCopy.createSheet(country1, 0);
	}

		@BeforeTest
		public void setUp() throws Exception
		{
			System.setProperty("webdriver.gecko.driver","C:/Users/PC/Desktop/Personal/selenium/Libarary/geckodriver-v0.33.0-win64/geckodriver.exe");
			driver = new FirefoxDriver();
		}
		
		public void SelectSearch() throws InterruptedException
		{
			
			Thread.sleep(1000);
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));			
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/header/div/nav[3]/ul/li[1]/a/span")));			
			driver.findElement(By.xpath("/html/body/div[1]/header/div/nav[3]/ul/li[1]/a/span")).click();
		}
		
		public void LoginPage()throws IOException,InterruptedException
		{
			driver.get("https://en.numista.com");
			
			driver.findElement(By.xpath("/html/body/div[1]/header/div/div[1]/div[5]/a[1]")).click();
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));			
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/div[1]/div/main/form/div/div[4]/input[3]")));
			
			driver.findElement(By.xpath("//*[@id=\"pseudo_connexion\"]")).sendKeys("Vijvijay1");
			driver.findElement(By.xpath("//*[@id=\"mdp_connexion\"]")).sendKeys("Dec@2009");
			driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form/div/div[4]/input[3]")).click();
			cururl=driver.getCurrentUrl();
		}
		
		
		
		void waitForLoad(WebDriver driver)
		{
			//new WebDriverWait(driver, Duration.ofSeconds(30)).until((ExpectedCondition<Boolean>) wd -> 
			new WebDriverWait(driver, Duration.ofSeconds(30)).until((ExpectedCondition<Boolean>) wd -> 
			((JavascriptExecutor)wd).executeScript("return document.readyState").equals("complete"));	
		}
		
		public void clickNextPage()throws IOException, RowsExceededException,WriteException, InterruptedException
		{
			if(z!=1)
			{
				driver.get(url1);;
			}
			else
			{
				driver.get(url);
			}
		
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
			WebElement until = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/footer/div/div/a[1]"))); // bottom page contact us xpath
			try
			{
				driver.findElement(By.linkText("Next")).click();
			}
			catch(Exception e)
			{
				e.printStackTrace();
			}
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/footer/div/div/a[1]")));
			url1 =driver.getCurrentUrl();
			z++;
		}
		
		public WritableCellFormat getfont(Colour color, int size)
		{
			WritableFont bold10font = new WritableFont(WritableFont.TAHOMA,size,WritableFont.BOLD,false,UnderlineStyle.NO_UNDERLINE,color);
			WritableCellFormat bold10format = new WritableCellFormat (bold10font);
			try
			{
				bold10format.setWrap(true);		
			}
			catch (WriteException e)
			{
				e.printStackTrace();
			}
			return bold10format;
		}
			
		public void coinInformation(String year, String comment,String C_VG,String C_F,String C_VF,String C_XF,String C_AU,String C_UNC)throws IOException, RowsExceededException,WriteException, InterruptedException
		{
			Label l1 = new Label(3, i, year);
			wshTemp.addCell(l1);
			//Label nil = new Label(4,i,"0");
			//wshTemp.addCell(nil);
			Label com = new Label(5,i,comment);
			wshTemp.addCell(com);
			
			Label C_VG1 = new Label(6,i,C_VG);
			wshTemp.addCell(C_VG1);
			
			Label C_F1 = new Label(7,i,C_F);
			wshTemp.addCell(C_F1);
			
			Label C_VF1 = new Label(8,i,C_VF);
			wshTemp.addCell(C_VF1);
			
			Label C_XF1 = new Label(9,i,C_XF);
			wshTemp.addCell(C_XF1);
			
			Label C_AU1 = new Label(10,i,C_AU);
			wshTemp.addCell(C_AU1);
			
			Label C_UNC1 = new Label(11,i,C_UNC);
			wshTemp.addCell(C_UNC1);
			
			i++;
			blank++;
			avail=false;
		}
		
		
		
		public void Select_100Page() throws IOException, RowsExceededException,WriteException,InterruptedException
		{
			//Click search
			
			driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[1]/input[2]")).click();
			
			//driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form/div/div[1]/input[2]")).click();
			
		
			
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
			//click on "display option"
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"results_options_switch\"]")));
			driver.findElement(By.xpath("//*[@id=\"results_options_switch\"]")).click();
			//click on 100
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/div/a[8]")));
			driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/div/a[8]")).click();			
			
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/footer/div/div/a[1]")));
			}
		
		public void Select_Country()throws IOException, RowsExceededException, WriteException, InterruptedException
		{
			driver.findElement(By.xpath("//*[@id='search_filters_button']")).click();
					
			WebElement dropdown3 = driver.findElement(By.id("select2-e-container"));
			dropdown3.click();
			Thread.sleep(1000);
			
			driver.findElement(By.xpath("/html/body/span/span/span[1]/input")).sendKeys(Cuntry);
		
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
			Thread.sleep(1000);
			
			WebElement dropdown1 = driver.findElement(By.id("select2-e-results"));
			List<WebElement> options = dropdown1.findElements(By.tagName("span"));
			for (WebElement option : options)
			{
				if(option.getText().contentEquals(Cuntry))
				{
					option.click();
					break;				
				}
			}
			
			// select metal type
			
			//Select sel = new Select(driver.findElement(By.id("m")));
			Select drpComposition = new Select(driver.findElement(By.name("m")));
			drpComposition.selectByVisibleText("Non precious");
			
			/*
			 * WebElement dropdown2 = driver.findElement(By.id("//*[@id=\\\"m\\"));
			 * List<WebElement> options1 = dropdown2.findElements(By.tagName("span")); for
			 * (WebElement option1 : options1) {
			 * if(option1.getText().contentEquals("Non precious")) { option1.click(); break;
			 * } }
			 */
			
			
			//driver.findElement(By.xpath("//*[@id=\"m\"]")).click();
			
			//Select sel = new Select(driver.findElement(By.id("//*[@id='m']")));
			
			
			//sel.deselectByVisibleText("Non precious");
			
			
			//select the coin type
			driver.findElement(By.xpath("//*[@id='coin_type_button']")).click();
						
				
			
				
				
				
			Boolean StdCirCoinisChecked  = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[2]/label/input")).isSelected();
			if (StdCirCoinisChecked.equals(false))
			{
				driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[2]/label/input")).click();
			}
			
			Boolean CommerCoinisChecked = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[3]/label/input")).isSelected();
			if (CommerCoinisChecked.equals(false))
			{
				driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[3]/label/input")).click();
			}
				
			Boolean NonCirCoinisChecked = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[4]/label/input")).isSelected();
			if (NonCirCoinisChecked.equals(true))
			{
				driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[4]/label/input")).click();
			}
				
			Boolean PatternCoinisChecked = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[5]/label/input")).isSelected();
			if (PatternCoinisChecked.equals(true))
			{
				driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[5]/label/input")).click();
			}	
				
			Boolean TokenCoinisChecked = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[6]/label/input")).isSelected();
			if (TokenCoinisChecked.equals(true))
			{
				driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/form[1]/div/div[4]/div[1]/div/ul/li[6]/label/input")).click();
			}	

			
		
			Select_100Page();
			
			try
			{
				driver.findElement(By.xpath(".//*[@id='resultats_recherche']/div[1]/a[2]")).isDisplayed();
				//List<WebElement> elements=driver.findElements(By.xpath("//p//a[contains(@href,'inex')]"));
				// Idendify the no of pages using /a and href 
				
				List<WebElement> elements=driver.findElements(By.xpath(".//a[contains(@href,'index.php?e=')]"));
				System.out.println("Total pages:"+(elements.size()));
				nop=(elements.size());
				String x = elements.get(nop-2).getText();
				nop = Integer.parseInt(x);
				String nowp = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/nav/a["+nop+"]")).getText();
				//String nowp = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/nav")).getText();
				if (nowp.contains("Next"))
					
				{
					nop = Integer.parseInt(driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/nav/a["+(nop-1)+"]")).getText());
				}
				else
				{
					nop = Integer.parseInt(driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div/div[2]/nav/a["+(nop)+"]")).getText());
				}
			
			}
			catch(Exception e)
			{
				nop=1;
			}
		}	
			
			
			
		public String CoinMetalFind()throws IOException, RowsExceededException, WriteException, InterruptedException
		{
			int Row_count = driver.findElements(By.xpath("/html/body/div[1]/div[1]/div/main/section[1]/table/tbody/tr/th")).size();
			String first_part = "/html/body/div[1]/div[1]/div/main/section[1]/table/tbody/tr[";
			String second_part1 = "]/th";
			String second_part2 = "]/td";
			//String third_part = "1]";
			
			for (int ix=1; ix<=Row_count;ix++)
			{
				String final_Header = first_part+ix+second_part1;
				String final_Headervalue = first_part+ix+second_part2;
				String Table_data_Header = driver.findElement(By.xpath(final_Header)).getText();
				String Table_Headervalue = driver.findElement(By.xpath(final_Headervalue)).getText();
				
				if (Table_data_Header.equals("Composition"))
				{
					Label Yrs1 = new Label(0,i+1,Table_Headervalue);
					return Yrs1.getString();
				}
			}return "NA";
		}	
			
		public void CoinDetails()throws IOException, RowsExceededException, WriteException, InterruptedException
		{
			int Row_count = driver.findElements(By.xpath("/html/body/div[1]/div[1]/div/main/section[1]/table/tbody/tr/th")).size();
			String first_part = "/html/body/div[1]/div[1]/div/main/section[1]/table/tbody/tr[";
			String second_part1 = "]/th";
			String second_part2 = "]/td";
			//String third_part = "1]";
			for ( int ix=1; ix<=Row_count;ix++)
			{
				String final_Header = first_part+ix+second_part1;
				String final_Headervalue = first_part+ix+second_part2;
				String Table_data_Header = driver.findElement(By.xpath(final_Header)).getText();
				String Table_Headervalue = driver.findElement(By.xpath(final_Headervalue)).getText();			
				if ( Table_data_Header.equals(("Years"))||(Table_data_Header.equals(("Year"))))
				{
					Label Yrs1 = new Label(0,i,Table_Headervalue);
					wshTemp.addCell(Yrs1);
				}
				if ( Table_data_Header.equals("Composition"))
				{
					Label Yrs1 = new Label(0,i+1,Table_Headervalue);
					wshTemp.addCell(Yrs1);
				}
				if ( Table_data_Header.equals("References"))
				{
					Label Yrs1 = new Label(0,i+2,Table_Headervalue);
					wshTemp.addCell(Yrs1);
				}
					
			}
		}
			
			
			
		/*################################*/
		/* TO FIND THE REQUIRED COINS */
		
		private class CoinBean
		{
			private String year;
			private String comment;
			private String C_VG;
			private String C_F;
			private String C_VF;
			private String C_XF;
			private String C_AU;
			private String C_UNC;
			
			public String getYear()
			{
				return year;
			}
			public void setYear(String year)
			{
				this.year = year;
			}
			public String getComment()
			{
				return comment;
			}
			public void setComment(String comment)
			{
				this.comment = comment;
			}
			
			public String getC_VG()
			{
				return C_VG;
			}
			public void setC_VG(String C_VG)
			{
				this.C_VG = C_VG;
			}
			
			
			public String getC_F()
			{
				return C_F;
			}
			public void setC_F(String C_F)
			{
				this.C_F = C_F;
			}
			
			public String getC_VF()
			{
				return C_VF;
			}
			public void setC_VF(String C_VF)
			{
				this.C_VF = C_VF;
			}
			
			public String getC_XF()
			{
				return C_XF;
			}
			public void setC_XF(String C_XF)
			{
				this.C_XF = C_XF;
			}
			
			
			public String getC_AU()
			{
				return C_AU;
			}
			public void setC_AU(String C_AU)
			{
				this.C_AU = C_AU;
			}
			
			public String getC_UNC()
			{
				return C_UNC;
			}
			public void setC_UNC(String C_UNC)
			{
				this.C_UNC = C_UNC;
			}
			
			public CoinBean(String year, String comment,String C_VG,String C_F,String C_VF,String C_XF,String C_AU,String C_UNC)
			{
				super();
				this.year = year;
				this.comment = comment;
				this.C_VG = C_VG;
				this.C_F = C_F;
				this.C_VF = C_VF;
				this.C_XF = C_XF;
				this.C_AU = C_AU;
				this.C_UNC = C_UNC;
			}
		}
			
		public void georgeAlogrithm()throws IOException, RowsExceededException, WriteException, InterruptedException
		{
			boolean toBeAdded=true;
			boolean isValueSet = false;
			String prevComment;
			ArrayList<CoinBean> coinList=new ArrayList<>();
			CoinBean coinBean=null;
			int runm2 = 0, runm3 = 3;
			runm2 =6;
			int td12 = driver.findElements(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td")).size();
			//currentYear=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[2]/tr/td[1]")).getText();
			currentYear=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[1]")).getText();
			
			try
			{			
				C_VG = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[4]")).getText();
				C_F  = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[5]")).getText();
				C_VF = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[6]")).getText();
				C_XF = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[7]")).getText();
				C_AU = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[8]")).getText();
				C_UNC =driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[9]")).getText();			
			}
			catch(Exception e)
			{
				System.out.println("No more coin values found");
			}
			
			
			if(td12 > 11)
			{
				prevComment=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody[6]/tr/td[13]")).getText();
			}
			else
			{
				
				prevComment = "";
			}
			int yearcount = driver.findElements(By.xpath("//td[contains(@class,'date')]")).size();
			
			for(rnum1=1;rnum1<=yearcount;rnum1++)
			{
				int Trow = driver.findElements(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody")).size();
				if (runm3 <=Trow)
				{
					Cctable = driver.findElements(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr")).size();
					foo = 0;
					foo1 = 0;
					for ( int CCcount = 1;CCcount<=Cctable;CCcount++)
					{
						try
						{
							//CCValue = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+(runm2+1)+"]/tr["+CCcount+"]/td[1]/span")).getText();
							CCValue = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+(runm2+1)+"]/tr["+CCcount+"]/td[1]/span")).getText();
							Full_CCValue = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+(runm2+1)+"]/tr["+CCcount+"]/td[1]")).getText();
												}
						//catch(NumberFormatException e)
						catch(Exception e)
						{
							foo = 0;
							CCValue = "0";
							Full_CCValue = "0";
						}
						//if (CCValue.endsWith("×")&&(!CCValue.contains("wap")) )
						if (CCValue.endsWith("×")&&(!Full_CCValue.contains("wap")) )
						{
							CCValue= CCValue.substring(0,CCValue.length()-1);
						}
					/*	if (CCValue.contains("wap") )
						{
							CCValue= CCValue.substring(0,CCValue.length()-1);
						}*/
						try {
							foo = Integer.parseInt(CCValue);
						}
						catch (NumberFormatException e)
						{
							foo=0;
						}
						foo1 = foo1+foo;
						
					}
					if (foo1==0)
					{
						s = null;
					}
					else
					{
						s = Integer.toString(foo1);
					}
				}
				try
				{
					yr=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[1]")).getText();
					C_VG_new = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[4]")).getText();
					C_F_new  = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[5]")).getText();
					C_VF_new = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[6]")).getText();
					C_XF_new = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[7]")).getText();
					C_AU_new = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[8]")).getText();
					C_UNC_new =driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[9]")).getText();
					
					if (td12 >11)
					{
						comment=driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[13]")).getText();
					}
					else
					{
						comment = "";
						
					}
				
				
				}
				catch(Exception e)
				{
					//if( !yr.contains("×"))
					//{
						yr ="0";
					//}
				}
				// to find coin market value
				/*
				 try
				{
					if (yr!="0")
					{
				C_VG = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[4]")).getText();
				C_F  = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[5]")).getText();
				C_VF = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[6]")).getText();
				C_XF = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[7]")).getText();
				C_AU = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[8]")).getText();
				C_UNC =driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/section[3]/table/tbody["+runm2+"]/tr/td[9]")).getText();
					}
				}
				catch(Exception e)
				{
					System.out.println("No more coin values found");
				}
				*/
				
				//if( !yr.contains("×"))
				//{
					if(!currentYear.contentEquals(yr) && (yr!="0"))
					{
						// will print/save below code only if there are only 1 year available for a KM ( if that coin is missing in our collection)
						if(toBeAdded && isValueSet)
						{
							System.out.println("Year "+currentYear+"Comment "+prevComment);
							coinBean=new CoinBean(currentYear, prevComment,C_VG,C_F,C_VF,C_XF,C_AU,C_UNC);
							coinList.add(coinBean);					
						}
						toBeAdded=true;
						isValueSet=false;
						prevComment = "";
						currentYear=yr;
						C_VG = C_VG_new;
						C_F = C_F_new;
						C_VF = C_VF_new;
						C_XF = C_XF_new;
						C_AU = C_AU_new;
						C_UNC = C_UNC_new;
					}
				//}
				if (s !=null && !s.isEmpty())
				{
					toBeAdded=false;
					isValueSet=true;				
				}
				else
				{
					if((comment.contains("proof"))||(comment.contains("Proof"))||(comment.contains("Sets"))||(comment.contains("sets")))
					{
						System.out.println("Year "+currentYear+"Comment "+prevComment);
					}
					else
					{
						prevComment=comment;
						isValueSet=true;					
					}
				}
				runm2 = runm2+2;
				//runm2 = runm2+1;
				runm3 = runm3+2;
				//runm3 = runm3+1;
			}
			// will print/save below code only if there are multiple years available for a KM ( if that coin is missing in our collection)
			if(toBeAdded && isValueSet)
			{
				System.out.println("Year "+currentYear+"Comment "+prevComment);
				coinBean=new CoinBean(currentYear, prevComment,C_VG,C_F,C_VF,C_XF,C_AU,C_UNC);
				coinList.add(coinBean);
			}
			
			generateExcel(coinList);
		}
		
		private void generateExcel(ArrayList<CoinBean>beanList) throws RowsExceededException, WriteException,IOException,InterruptedException
		{
			for(CoinBean bean: beanList)
			{
				if(ImageCapture.equals(false))
				{
					CaptureCoinImage();
					CoinDetails();
				}
				coinInformation(bean.getYear(),bean.getComment(),bean.getC_VG(),bean.getC_F(),bean.getC_VF(),bean.getC_XF(),bean.getC_AU(),bean.getC_UNC());
			}
		}
		
		
		public void CaptureCoinImage()throws IOException,RowsExceededException,WriteException, InterruptedException
		{
			Thread.sleep(100);
			String Title = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/header/h1")).getText();
			
			try 
			{
				Label labTemp = new Label(0,i++,Title,getfont(Colour.BLUE,10));
				wshTemp.addCell(labTemp);
			}
			catch(Exception e)
			{
				System.out.println("Error writing in Excelsheet");
			}
			
			
			for ( int a=1;a<=2;a++)
			{
				Thread.sleep(100);
				WebElement ele= null;
				try
				{
					ele = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div/main/div[2]/a["+a+"]/img"));
					Thread.sleep(100);
					TakeScreenShot shot=new TakeScreenShot(driver);
					File screenShot = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					screenShot = shot.elelocation(ele);
				
					FileUtils.copyFile(screenShot, new File(System.getProperty("user.dir")+"\\"+"screenshot1.png"));
					File imageFile = new File(System.getProperty("user.dir")+"\\"+"screenshot1.png");
					BufferedImage input = ImageIO.read(imageFile);
					ByteArrayOutputStream baos = new ByteArrayOutputStream();
					ImageIO.write(input,"PNG",baos);
					wshTemp.addImage(new WritableImage(q,i,1,3,baos.toByteArray()));
					index++;
					ImageCapture=true;
						if ( count%2==1)
							q=q+1;
						count=count+1;
				}
				catch(Exception e)
				{
					System.out.println("ImageNotAvailable");
					ImageCapture=true;
					if(count%2==1)
						q=q+1;
					count=count+1;
				}
			}
		}
		
			public void MissingCoinsCapture()throws IOException,InterruptedException, RowsExceededException, WriteException
			{
				int PgNo = Integer.parseInt(PgNo1);
				for ( int pageno=1;pageno<=nop;pageno++)
				{
				
				WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/footer/div/div/a[1]")));
				List<WebElement>elements = driver.findElements(By.xpath("//div[contains(@class,'resultat_recherche')]//a"));
				int u=0;
				String[] links =new String[(elements.size()/3)];
				for (int b=0;b<elements.size();b=b+3)
				{
					WebElement e=elements.get(b);
					links[u]=e.getAttribute("href");
					u++;
				}
				int nolinks=u;
				if(pageno >=PgNo)
				{
					for (u=0;u<nolinks;u++)
					{
						driver.get(links[u]);
						wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/footer/div/div/a[1]")));
						
						WebElement table = driver.findElement(By.className("collection"));
						List<WebElement>allRows = table.findElements(By.tagName("tr"));
						nrow=allRows.size();
						ImageCapture = false;
						blank=0;
						String Metal_Type = CoinMetalFind();
						try 
						{
							if ( Metal_Type.contentEquals(null))
							{
								Metal_Type="NA";
							}		
						}
						catch(Exception e)
						{
							System.out.println("Error CoinMetal");
						}
					
						if(((Metal_Type!=null)&&(!(Metal_Type.contains("Gold")))&&(!(Metal_Type.contains("Silver")))&&(!(Metal_Type.contains("Platium"))))||(Metal_Type.contains("plated"))||(Metal_Type.contains("Nordic")))
						{
							georgeAlogrithm();
							if ((blank<=3)&&(ImageCapture.equals(true)))
							{
								i=i+2;
							}
							avail=false;
									q=1;
							Thread.sleep(100);
						}
						
					}
				}
				if(pageno!=nop)
				{
					clickNextPage();
				}
			
		}	
					
					
	}
}		
				
				
		
	
	
	
	
	
	
	
	
	












