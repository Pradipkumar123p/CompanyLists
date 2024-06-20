package CompanyList;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Company_List1 {
	
	public static void main(String[] args) throws IOException, InterruptedException {
		
		 String excelFilePath = "C:\\Users\\nizam\\Downloads\\CompanyListEurope.xlsx";
	        String csvFilePath = "C:\\Users\\nizam\\eclipse-workspace\\Java\\Excel_File\\ListCompany_Europe.csv"; // Adjust path as per your environment

	        System.setProperty("webdriver.chrome.driver", "C:\\Users\\nizam\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

	        ChromeOptions op = new ChromeOptions();
	        op.addArguments("--remote-allow-origins=*","headless");
	        op.setBinary("C:\\Users\\nizam\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe");

	        WebDriver driver = new ChromeDriver(op);
	        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));

	        try {
	            driver.get("https://www.google.com/");

	            // Example search
	            String searchKeyword = "build with";
	            driver.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys(searchKeyword);
	            driver.findElement(By.xpath("//textarea[@class='gLFyf']")).sendKeys(Keys.ENTER);
	            driver.findElement(By.xpath("(//h3[@class='LC20lb MBeuO DKV0Md'])[1]")).click();

	          //  driver.findElement(By.xpath("(//h3)[1]")).click(); // Click the first search result

	            FileInputStream file = new FileInputStream(excelFilePath);
	            XSSFWorkbook workbook = new XSSFWorkbook(file);
	            XSSFSheet sheet = workbook.getSheetAt(0);
	          
	            int rowCount = sheet.getLastRowNum();
	            System.out.println(rowCount);

	            List<String[]> resultList = new ArrayList<>();
	            resultList.add(new String[]{"Company URL", "Company_Status", "Platform_Status"}); // Header for CSV

	            for (int i = 1; i <= 5; i++) {
	                Row row = sheet.getRow(i);
	                Cell cell = row.getCell(2); // Assuming URL is in column C (index 2)
	                String cellvalue = cell.getStringCellValue();
	                System.out.println(cellvalue);

	                // Your logic to fetch company status and platform status
	                String companyURL = cellvalue;
	                String companyStatus = ""; // Replace with your logic
	                String platformStatus = "Unknown"; // Replace with your logic
	                
	                
	                driver.findElement(By.xpath("//input[@type='text']")).sendKeys(companyURL);
	                driver.findElement(By.xpath("//input[@type='submit']")).click();
	                

	                Thread.sleep(1000);
	                List<WebElement> textMatche = driver.findElements(By.xpath("//p"));
	                List<WebElement> textMatches1 = driver.findElements(By.xpath("//a[@class='text-dark']"));

	                ArrayList<WebElement> textMatches = new ArrayList<WebElement>();
	                textMatches.addAll(textMatche);
	                textMatches.addAll(textMatches1);

	                String platform = "Not add platform";
	                String matchText = "Not Match";
	                for (WebElement matchElement : textMatches) {
	                    String text = matchElement.getText().toLowerCase();
	                     if (text.contains("ecommerce")) {
	                        platform = "E-commerce";
	                    }
	                     else if (text.contains("shopify")) {
	                        matchText = "Shopify";
	                        break;
	                    }
	                    else if (text.contains("weebly")) {
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
	                    } 
	                }

	                System.out.println(matchText);
	                System.out.println(platform);
	                
	                driver.navigate().back();
	                Thread.sleep(1000);
	                driver.findElement(By.xpath("//input[@type='text']")).clear();

	                // Example of collecting data to write to CSV
	                resultList.add(new String[]{companyURL, companyStatus, platformStatus});
	                
	                
	            }

	            file.close();

	            // Write resultList to CSV file
	            writeResultsToCSV(csvFilePath, resultList);
	            System.out.println("CSV file created successfully at: " + csvFilePath);

	        } finally {
	            driver.quit();
	        }
	    }


	    static void writeResultsToCSV(String csvFile, List<String[]> resultList) throws IOException {
	        File file = new File(csvFile);
	        file.getParentFile().mkdirs(); // Ensure parent directories exist
	        try (CSVPrinter printer = new CSVPrinter(new FileWriter(file), CSVFormat.DEFAULT)) {
	            for (String[] result : resultList) {
	                printer.printRecord((Object[]) result);
	            }
	        }
	    }
	                                                                                
    }
	            
       
	
   
     
































/*  public static String getcellvalue(XSSFCell cell)    {
	  
	  switch (cell.getCellType())   {
    	case NUMERIC:
		
		return String.valueOf(cell.getNumericCellValue());
		
        case BOOLEAN:
		
		return String.valueOf(cell.getBooleanCellValue());
		
         case STRING:
		
		 return cell.getStringCellValue();
		
		default:
			
			return cell.getStringCellValue();
	
	}
	  
	  
  }   */
	

