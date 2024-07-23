package CompanyList;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ListOfCompanyForSouthAmrica {
	

	  public static void main(String[] args) throws IOException, InterruptedException {
	    	
	        String filepath = "C:\\Users\\nizam\\eclipse-workspace\\OfficeWork\\Excel_File\\";
	        String filename = "CompanyList_SouthAmrica.xlsx";

	        System.setProperty("webdriver.chrome.driver", "C:\\Users\\nizam\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

	        ChromeOptions op = new ChromeOptions();
	        op.addArguments("--remote-allow-origins=*","headless");
	        op.setBinary("C:\\Users\\nizam\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe");

	        WebDriver driver = new ChromeDriver(op);
	        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
	        driver.get("https://www.google.com/");
	        driver.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys("build with");
	        driver.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys(Keys.ENTER);
	        driver.findElement(By.xpath("(//h3[@class='LC20lb MBeuO DKV0Md'])[1]")).click();

	        FileInputStream file = new FileInputStream("C:\\Users\\nizam\\Downloads\\CompanyListSouthAmrica.xlsx");
	        XSSFWorkbook book = new XSSFWorkbook(file);
	        XSSFSheet sheet = book.getSheetAt(0);
	        int rowCount = sheet.getLastRowNum();
	        System.out.println(rowCount);

	        // Create the new workbook and sheet once outside the loop
	        XSSFWorkbook book1 = new XSSFWorkbook();
	        XSSFSheet sheet1 = book1.createSheet("Sheet1.1");

	        // Create header style
	        XSSFCellStyle headerStyle = book1.createCellStyle();
	        headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
	        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	        // Create header row
	        XSSFRow headerRow = sheet1.createRow(0);
	        String[] headers = {"Company_URL", "Company_Status", "Platform_Status"};
	        for (int i = 0; i < headers.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(headers[i]);
	            cell.setCellStyle(headerStyle); // Apply style to header cells
	        }

	        for (int i = 2; i <= rowCount; i++) {
	            Row r = sheet.getRow(i);
	            Cell c = r.getCell(2);
	            String cellValue = c.getStringCellValue();

	            System.out.println(cellValue);
	            Thread.sleep(500);
	            driver.findElement(By.xpath("//input[@type='text']")).sendKeys(cellValue);
	            driver.findElement(By.xpath("//input[@type='submit']")).click();

	            Thread.sleep(2500);
	            List<WebElement> textMatche = driver.findElements(By.xpath("//p"));
	            List<WebElement> textMatches1 = driver.findElements(By.xpath("//a[@class='text-dark']"));
	            List<WebElement> textMatches2 = driver.findElements(By.xpath("//h6[@class='card-title text-secondary']"));

	            ArrayList<WebElement> textMatches = new ArrayList<>();
	            textMatches.addAll(textMatche);
	            textMatches.addAll(textMatches1);
	            textMatches.addAll(textMatches2);

	            String platform = "Not add platform";
	            String matchText = "Not Match";
	            for (WebElement matchElement : textMatches) {
	                String text = matchElement.getText().toLowerCase();
	                if (text.contains("ecommerce")) {
	                    platform = "E-commerce";
	                } else if (text.contains("shopify")) {
	                    matchText = "Shopify";
	                    break;
	                } else if (text.contains("weebly")) {
	                    matchText = "Weebly";
	                    break;
	                } else if (text.contains("woocommerce")) {
	                    matchText = "WooCommerce";
	                    break;
	                } else if (text.contains("bigcommerce")) {
	                    matchText = "BigCommerce";
	                    break;
	                } else if (text.contains("wordpress")) {
	                    matchText = "WordPress";
	                    break;
	                } else if (text.contains("salesforce commerce cloud")) {
	                    matchText = "Salesforce commerce cloud";
	                    break;
	                } else if (text.contains("magento")) {
	                    matchText = "Magento";
	                    break;
	                } else if (text.contains("wix")) {
	                    matchText = "Wix";
	                    break;
	                } else if (text.contains("prestashop")) {
                      matchText = "PrestaShop";
                      break;
                  } 
	            }

	            System.out.println(matchText);
	            System.out.println(platform);
	            driver.navigate().back();
	            Thread.sleep(1000);
	            driver.findElement(By.xpath("//input[@type='text']")).clear();

	            System.out.println("Before Create List: " + isFileExist(filepath + filename));

	            // Add a row to the sheet in each iteration
	            XSSFRow rowList = sheet1.createRow(i - 1); // Adjust the index to start from 0
	            rowList.createCell(0).setCellValue(cellValue);
	            rowList.createCell(1).setCellValue(matchText);
	            rowList.createCell(2).setCellValue(platform);

	            System.out.println("After Create List: " + isFileExist(filepath + filename));
	        }

	        // Auto size columns for the new sheet
	        for (int i = 0; i < headers.length; i++) { // Adjust to fit all columns
	            sheet1.autoSizeColumn(i);
	        }

	        // Write to the file once after the loop
	        try (FileOutputStream fs = new FileOutputStream(filepath + filename)) {
	            book1.write(fs);
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (Exception e) {
	            e.printStackTrace();
	        } finally {
	            book1.close();
	            file.close();
	            driver.quit();
	        }
	    }

	    static boolean isFileExist(String filepath) {
	        File f = new File(filepath);
	        return f.exists();
	    }
	
	
	

}
