package PerksData;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ExtractData_Termscondition {
	
	 private static final int EXCEL_CELL_CHARACTER_LIMIT = 32767;

    public static void main(String[] args) throws InterruptedException, IOException {
    	
        String filepath = "C:\\Users\\nizam\\eclipse-workspace\\OfficeWork\\Excel_File\\";
        String filename = "Terms&Condition_List.xlsx";

        System.setProperty("webdriver.chrome.driver", "C:\\Users\\nizam\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        op.setBinary("C:\\Users\\nizam\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe");

        WebDriver driver = new ChromeDriver(op);

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));

        List<String> urls = readUrlsFromExcel("C:\\Users\\nizam\\Downloads\\exibitor_privacy_terms.xlsx");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Extract Data");

        Row header = sheet.createRow(0);
        Cell headerCell1 = header.createCell(0);
        headerCell1.setCellValue("URL");
        Cell headerCell2 = header.createCell(1);
        headerCell2.setCellValue("Terms and Condition & privacy Policy");
        
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
       
        
        headerCell1.setCellStyle(headerStyle);
        headerCell2.setCellStyle(headerStyle);

        int rowIndex = 1;
        
        
        for (String url : urls) {
            driver.get(url);
            
            String content = "";

            switch (url) {
            case "https://www.enyamond.com/policies/privacy-policy":
            	
            content = driver.findElement(By.xpath("//div[@class='main-container container']")).getText();
            System.out.println(content);
            break;

            case "https://www.enyamond.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.biogaia.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='entry-content']")).getText();
                System.out.println(content);
                break;

            case "https://alevanaturals.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.biogaia.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='entry-content']")).getText();
                System.out.println(content);
                break;

            case "https://bujjify.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://bujjify.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://bujjify.com/policies/shipping-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://cheekybambino.ca/pages/return-and-exchange-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='page-description rich-editor-text-content']")).getText();
                System.out.println(content);
                break;

            case "https://www.cefi.ca/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@id='_rich_text-5-3']")).getText();
                System.out.println(content);
                break;

     /*       case "https://www.costco.ca/customerservice.costco.ca/app/answers/answer_view/a_id/585":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;                             */

            case "https://www.costco.ca/purchase-terms.html":
            	
            	content = driver.findElement(By.xpath("//div[@class='espot']")).getText();
                System.out.println(content);
                break;

            case "https://www.costco.ca/terms-and-conditions-of-use.html":
            	
            	content = driver.findElement(By.xpath("//p[@class='tcCase']")).getText();
                System.out.println(content);
                break;

            case "https://www.costco.ca/privacy-policy.html":                         //148
            	
            	content = driver.findElement(By.xpath("//div[@class='innerBlock']")).getText();
                System.out.println(content);
                break;

            case "https://alevanaturals.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://alevanaturals.com/policies/refund-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://alevanaturals.com/policies/shipping-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.drbrownsbaby.com/cookie-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='column ten offset-three']")).getText();  //174
                System.out.println(content);
                break;

            case "https://www.drbrownsbaby.com/terms-of-use/":
            	
            	content = driver.findElement(By.xpath("//div[@class='column ten offset-three']")).getText();
                System.out.println(content);
                break;

            case "https://hydralyte.com/pages/shipping-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte rtnu-page__content']")).getText();
                System.out.println(content);
                break;

            case "https://hydralyte.com/pages/refund-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte rtnu-page__content']")).getText();
                System.out.println(content);
                break;

            case "https://hydralyte.com/pages/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='elm text-edit']")).getText();
                System.out.println(content);
                break;

            case "https://hydralyte.com/pages/terms-conditions":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte rtnu-page__content']")).getText();
                System.out.println(content);
                break;

            case "https://www.insception.com/insception-pricing/refund-policy/":
            	
            	content = driver.findElement(By.xpath("//section[@class='post_content clearfix']")).getText();
                System.out.println(content);
                break;

            case "https://cellsforlife.com/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='et_pb_text_inner']")).getText();
                System.out.println(content);
                break;

            case "https://www.royale.ca/diapers/privacy-policy/":     //220
            	
            	content = driver.findElement(By.xpath("//h1[@class='css-1w7ejuy e13xswbp1']")).getText();
                System.out.println(content);
                break;

            case "https://www.royale.ca/diapers/terms-of-service/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://flektoys.com/en/en/pages/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://flektoys.com/en/en/pages/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://lotusbabyco.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://lotusbabyco.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.lumehra.com/#":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://makemybellyfit.com/pages/return-policy-politique-de-retour":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://makemybellyfit.com/pages/terms-conditions-termes-conditions":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://makemybellyfit.com/pages/privacy-policy-politique-de-confidentialite":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.ontario.ca/page/terms-use":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.munchkin.com/terms-conditions":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://www.munchkin.com/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://orangenaturals.com/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;

            case "https://orangenaturals.com/shipping-returns-exchanges":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                    
            case "https://orangenaturals.com/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://mypaume.com/pages/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://mypaume.com/pages/paume-in-terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://www.pegperego.com/en_ca/baby/privacy-policy-cookie-restriction-mode/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
            
            case "https://seaford.ca/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://prenabelt.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://talkinsleep.com/terms-and-conditions-of-use/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://talkinsleep.com/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://trippinalongboutique.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://trippinalongboutique.com/policies/refund-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://trippinalongboutique.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://uppababy.com/ordering-info-and-returns-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://uppababy.com/privacy-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://uppababy.com/terms-of-service/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://uppababy.com/cookie-policy/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://www.kenvue.com/privacy-policy/canada/en":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.com/our-policies/policy-on-privacy-of-customer-personal-information/":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.ca/policies/refund-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.ca/pages/chemical-sustainability-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.ca/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.ca/policies/shipping-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://clekinc.ca/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://caley-beth.com/policies/privacy-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://caley-beth.com/policies/shipping-policy":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            case "https://caley-beth.com/policies/terms-of-service":
            	
            	content = driver.findElement(By.xpath("//div[@class='rte']")).getText();
                System.out.println(content);
                break;
                
            default:
                System.out.println("No specific actions for this URL: " + url);
                break;
            }
            
            Row row = sheet.createRow(rowIndex++);
            Cell urlCell = row.createCell(0);
            urlCell.setCellValue(url);

            writeLargeTextToCell(sheet, rowIndex - 1, 1, content);

            System.out.println("Extracted content from URL: " + url);
        }

        for (int i = 0; i < 2; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream outputStream = new FileOutputStream(filepath + filename)) {
            workbook.write(outputStream);
        }

        workbook.close();
        driver.quit();
    }

    
    private static List<String> readUrlsFromExcel(String filePath) throws IOException {
    	
        List<String> urls = new ArrayList<>();
        
        try (FileInputStream inputStream = new FileInputStream(filePath);
        		
             Workbook workbook = new XSSFWorkbook(inputStream)) {
             Sheet sheet = workbook.getSheetAt(0);
            
            for (Row row : sheet) {
            	
                Cell cell = row.getCell(0);
                
                if (cell != null) {
                	
                    urls.add(cell.getStringCellValue());
                }
            }
        }
        
        return urls;
    }

    private static void writeLargeTextToCell(Sheet sheet, int rowIndex, int columnIndex, String text) {
    	
        String[] lines = text.split("\n");
        for (String line : lines) {
        	
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(columnIndex);
            cell.setCellValue(line);
            rowIndex++;
          }
      }
  }