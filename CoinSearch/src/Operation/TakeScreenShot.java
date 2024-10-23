package Operation;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class TakeScreenShot
{
	WebDriver driver;
public TakeScreenShot  (WebDriver driver)
{
		this.driver=driver;
}
	


public File elelocation(WebElement ele) throws IOException
{
	File screenShot = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	BufferedImage fullImage = ImageIO.read(screenShot);
	//Get the location on the page
	Point point = ele.getLocation();
	//Get width and height of an element
	int eleWidth = ele.getSize().getWidth();
	
	int eleHeight = ele.getSize().getHeight();
	//Cropping the entire page screen shot to have only element screenshot
	if(eleWidth != 0 && eleHeight !=0)
	{
		BufferedImage eleScreenShot = fullImage.getSubimage(point.getX(), point.getY(), eleWidth, eleHeight);
		ImageIO.write(eleScreenShot, "png", screenShot);
	}
	return screenShot;
	}
}

