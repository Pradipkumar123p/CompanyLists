package PerksData;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ExtractData_Speaker {


    public static void main(String[] args) throws InterruptedException, IOException {
    	
        String filepath = "C:\\Users\\nizam\\eclipse-workspace\\OfficeWork\\Excel_File\\";
	    String filename = "Speaker_List1.xlsx";

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\nizam\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        op.setBinary("C:\\Users\\nizam\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe");

        WebDriver driver = new ChromeDriver(op);

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
        Thread.sleep(500);

        driver.get("https://thebabyshows.com/toronto-fall-baby-show/#speakers");
        
        driver.manage().window().maximize();

        System.out.println(driver.getCurrentUrl());

        try {
            driver.findElement(By.xpath("(//button[@class='ub-emb-close'])[1]")).click();
        } catch (Exception e) {
            System.out.println("popup is not present");
        }

        List<WebElement> list = driver.findElements(By.xpath("//h2[@class='ab-profile-name']"));
        List<WebElement> list1 = driver.findElements(By.xpath("//p[@class='ab-profile-title']"));
        List<WebElement> list3 = driver.findElements(By.xpath("//div[@class='wp-block-button']//a[contains(text(),'FULL BIO')]"));
        List<WebElement> list4 = driver.findElements(By.xpath("//figure[@class='ab-profile-image-square']//img")); 

        System.out.println(list.size());
        System.out.println(list1.size());
        System.out.println(list3.size());
        System.out.println(list4.size());
        
        XSSFWorkbook book1 = new XSSFWorkbook();
        XSSFSheet sheet1 = book1.createSheet("Sheet1.1");

        // Create header style
        XSSFCellStyle headerStyle = book1.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFRow headerRow = sheet1.createRow(0);
        String[] headers = {"Speaker_Name", "Speaker_Title", "Speaker_SocialHandle", "Speaker_Image", "Speaker_Profile"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle); // Apply style to header cells
        }

        for (int i = 0; i < list.size(); i++) {
        	
            try {
            	Thread.sleep(500);
            	
                // Fetch the company name after clicking
                String SpeakerName = list.get(i).getText();
                System.out.println("Speaker Name: " + SpeakerName);

                String SpeakerTitle = list1.get(i).getText();
                System.out.println("Speaker Title: " + SpeakerTitle);
                
                String Image = list4.get(i).getAttribute("src").toString();
                System.out.println("SpeakerImage: " + Image);
                
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].click()", list3.get(i));
                Thread.sleep(500);
                
                WebElement text = driver.findElement(By.xpath("//div[@class='ab-profile-text']"));
                
                String SpeakerProfile = text.getText();
                System.out.println("Speaker Profile: " + SpeakerProfile);
               
                List<WebElement> socialHandles = driver.findElements(By.xpath("//ul[@class='ab-social-links']//li//a"));
                
                System.out.println(socialHandles.size());

                List<String> socialHandleList = new ArrayList<>();
                for (WebElement handle : socialHandles) {
                    socialHandleList.add(handle.getAttribute("href"));
                }
                String allSocialHandles = String.join(", ", socialHandleList);
                System.out.println("Speaker Social Handles: " + allSocialHandles);
                
                System.out.println("-----------------------");
                
                driver.navigate().back();
                
                // Add a row to the sheet in each iteration
                XSSFRow rowList = sheet1.createRow(i + 1); // Adjust the index to start from 1
                rowList.createCell(0).setCellValue(SpeakerName);
                rowList.createCell(1).setCellValue(SpeakerTitle);
                rowList.createCell(2).setCellValue(allSocialHandles); // Add phone number if available
                rowList.createCell(3).setCellValue(Image);
                rowList.createCell(4).setCellValue(SpeakerProfile);
            
                // Refresh the list of elements
               list = driver.findElements(By.xpath("//h2[@class='ab-profile-name']"));
               list1 = driver.findElements(By.xpath("//p[@class='ab-profile-title']"));
               socialHandles = driver.findElements(By.xpath("//ul[@class='ab-social-links']//li//a"));
               list3 = driver.findElements(By.xpath("//div[@class='wp-block-button']//a[contains(text(),'FULL BIO')]"));
               list4 = driver.findElements(By.xpath("//figure[@class='ab-profile-image-square']//img")); 

            } catch (Exception e) {
            	
                System.out.println("Failed to click element at index: " + i);
           }
         
	        
	        //Refresh the list of elements
             list = driver.findElements(By.xpath("//h2[@class='ab-profile-name']"));
             list1 = driver.findElements(By.xpath("//p[@class='ab-profile-title']"));
             list3 = driver.findElements(By.xpath("//div[@class='wp-block-button']//a[contains(text(),'FULL BIO')]"));
             list4 = driver.findElements(By.xpath("//figure[@class='ab-profile-image-square']//img")); 
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
