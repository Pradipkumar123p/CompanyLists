package PerksData;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.WindowType;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ExtractData_Exibitor1 {

    public static void main(String[] args) throws InterruptedException, IOException {
    	
        String filepath = "C:\\Users\\nizam\\eclipse-workspace\\OfficeWork\\Excel_File\\";
	    String filename = "Exibitor_List.xlsx";

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\nizam\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        op.setBinary("C:\\Users\\nizam\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe");

        WebDriver driver = new ChromeDriver(op);

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
        Thread.sleep(5000);

        driver.get("https://thebabyshows.com/toronto-fall-baby-show/");
        
        driver.manage().window().maximize();

        System.out.println(driver.getCurrentUrl());

        try {
            driver.findElement(By.xpath("(//button[@class='ub-emb-close'])[1]")).click();
        } catch (Exception e) {
            System.out.println("popup is not present");
        }

        List<WebElement> list = driver.findElements(By.xpath("//span[@class='uww_exh_list_list_compdrilldown']//a[@href='#']"));

        System.out.println(list.size());

        XSSFWorkbook book1 = new XSSFWorkbook();
        XSSFSheet sheet1 = book1.createSheet("Sheet1.1");

        // Create header style
        XSSFCellStyle headerStyle = book1.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFRow headerRow = sheet1.createRow(0);
        String[] headers = {"Company_Name", "Company_Address", "Company_PhoneNumber", "Company_Website"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle); // Apply style to header cells
        }

        for (int i = 0; i < list.size(); i++) {
            try {
                Thread.sleep(4000);
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].click()", list.get(i));
                Thread.sleep(4000); 

                // Fetch the company name after clicking
                WebElement companyNameElement = driver.findElement(By.id("uww_exh_popup_header")); // Replace with the correct XPath
                String companyName = companyNameElement.getText();
                System.out.println("Company Name: " + companyName);

                WebElement add = driver.findElement(By.xpath("(//div[@style='margin:10px;'])[2]"));
                String address = add.getText();

                // Validate if the address starts with an alphanumeric character
                if (address.matches("^[a-zA-Z0-9].*")) {
                    System.out.println("Address: " + address);
                } else {
                    address = "Address Not Found";
                }
                
                WebElement phone = driver.findElement(By.xpath("(//div[@style='margin:10px;'])[3]//a"));
                String phoneNumber = phone.getText();
                
                if (phoneNumber.startsWith("+1")) {
                    System.out.println("Phone Number: " + phoneNumber);
                } else {
                    phoneNumber = "Phone Number Not Found";
                }
                
                WebElement w2 = driver.findElement(By.xpath("(//div[@style='margin:10px;'])[4]//a"));
                String website = w2.getAttribute("href");
                System.out.println(website);

                System.out.println("-----------------------");
                
                // Add a row to the sheet in each iteration
                XSSFRow rowList = sheet1.createRow(i + 1); // Adjust the index to start from 1
                rowList.createCell(0).setCellValue(companyName);
                rowList.createCell(1).setCellValue(address);
                rowList.createCell(2).setCellValue(phoneNumber); // Add phone number if available
                rowList.createCell(3).setCellValue(website);
           
                // Close the popup or navigate back
                WebElement closeButton = driver.findElement(By.xpath("//i[@class='fa fa-times']"));
                js.executeScript("arguments[0].click()", closeButton);
                Thread.sleep(2000);

                // Refresh the list of elements
                list = driver.findElements(By.xpath("//span[@class='uww_exh_list_list_compdrilldown']//a[@href='#']"));

                // Open a new tab and perform actions
                WebDriver newTab = driver.switchTo().newWindow(WindowType.TAB);
                newTab.get("https://www.google.co.in/");
                
                newTab.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys("chatgpt");
                newTab.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys(Keys.ENTER);
                
                newTab.findElement(By.xpath("(//h3[@class='LC20lb MBeuO DKV0Md'])[2]")).click();
                
                newTab.findElement(By.xpath("//div[@class='flex flex-col items-start m:items-center m:flex-row gap-xs items-center']//a")).click();
                
                driver.switchTo().window(driver.getWindowHandles().iterator().next());  
                
                for(int i1 =0; i1<= list.size()-1; i1++)  {
                	
                newTab.findElement(By.xpath("//textarea[@placeholder='Message ChatGPT']")).sendKeys("give me short description for this" + website);
            
                newTab.findElement(By.xpath("(//div//button)[17]")).click();
                
                driver.switchTo().window(driver.getWindowHandles().iterator().next());  
                
                // Close the new tab and switch back to the original tab
                newTab.close();
                
                }
                
                driver.switchTo().window(driver.getWindowHandles().iterator().next());
                
                driver.close();
                
            } catch (Exception e) {
                System.out.println("Failed to click element at index: " + i);
            }
        }

        for (int i = 0; i < headers.length; i++) { // Adjust to fit all columns
            sheet1.autoSizeColumn(i);
        }

        // Write to the file once after the loop
        try (FileOutputStream fs = new FileOutputStream(filepath + filename)) {
            book1.write(fs);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            book1.close();
            driver.quit();
        }
    }

    static boolean isFileExist(String filepath) {
        File f = new File(filepath);
        return f.exists();
    }
}
