package ReadExcelFile;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.util.List;

public class SearchOnGoogle {
    public static void main(String[] args) throws InterruptedException {
        // Chrome Driver সেটআপ করো (Ensure chromedriver.exe আছে)
        //System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();

        // Google-এ প্রবেশ করো
        driver.get("https://www.google.com");

        // সার্চ বক্স খুঁজে "Dhaka" লিখে ENTER চাপো
        WebElement searchBox = driver.findElement(By.xpath("//textarea[@id='APjFqb']"));
        searchBox.sendKeys("Dhaka");
        Thread.sleep(1000);
        searchBox.sendKeys(Keys.ENTER);

        // Auto-suggestion লিস্ট বের করো
        List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']//li"));

        String longest = "", shortest = suggestions.get(0).getText();

        for (WebElement suggestion : suggestions) {
            String text = suggestion.getText();
            if (text.length() > longest.length()) {
                longest = text;
            }
            if (text.length() < shortest.length() && text.length() > 0) {
                shortest = text;
            }
        }

        System.out.println("Longest: " + longest);
        System.out.println("Shortest: " + shortest);

        // ডাটা Excel ফাইলে লিখতে পাঠাও
        Object ExcelWriter = null;
        ExcelWriter.wait();

        // ব্রাউজার বন্ধ করো
        driver.quit();
    }
}
