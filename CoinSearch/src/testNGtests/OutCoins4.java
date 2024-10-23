package testNGtests;

import java.io.File;
import jxl.read.biff.BiffException;
import java.io.IOException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import org.openqa.selenium.By;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class OutCoins4 extends Libary
{
@Test(dataProvider="testdata")
public void captureCoin(String no, String country1, String PgNo, String KM)throws IOException, RowsExceededException, WriteException, BiffException, InterruptedException
{
	LoginPage();
	createExcelCountry(country1);
	Cuntry = country1;
	PgNo1 = PgNo;
	SelectSearch();
	Select_Country();
	url=driver.getCurrentUrl();
	MissingCoinsCapture();
	wwbCopy.write();
	wwbCopy.close();
	
	
}

	@DataProvider(name="testdata")
	public Object [] [] readDriver() throws BiffException, IOException, RowsExceededException, WriteException
	{
		File f = new File(System.getProperty("user.dir")+"\\Driver\\"+"Driver.xls");
		Workbook w = Workbook.getWorkbook(f);
		Sheet s = w.getSheet("Sheet1");
		int rows = s.getRows();
		int columns = s.getColumns();
		String inputData[][] = new String[rows-1][columns];
		int i=1;
		for (i=1;i<rows;i++)
			{
			for(int j=0;j<columns;j++)
			{
			Cell c = s.getCell(j,1);
			inputData[i-1][j] = c.getContents();
			}
			}
	return inputData;
}

}




