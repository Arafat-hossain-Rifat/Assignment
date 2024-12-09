package test;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

import java.time.Duration;

public class BaseTest {

    protected final String URL="https://www.google.com/";
    protected WebDriver driver;
    protected final int MAX_WAIT=10;
    protected WebDriverWait wait;

    @BeforeMethod
    public void setup(){
        driver=new ChromeDriver();
        driver.manage().window().maximize();
        driver.get(URL);
        wait=new WebDriverWait(driver, Duration.ofSeconds(MAX_WAIT));
    }

    @AfterMethod
    public void tearDown(){
        driver.quit();
    }
}
