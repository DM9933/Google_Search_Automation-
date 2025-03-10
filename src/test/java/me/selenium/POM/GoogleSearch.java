package me.selenium.POM;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class GoogleSearch{

    WebDriver driver;
    String excelFilePath = "D:\\IT Training BD\\Java\\Test_Automation_Q1_Assignments\\Excel.xlsx";

    @BeforeSuite
    public void openBrowser() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
    }

    @Test
    public void googleSearchTest() throws InterruptedException {
        String today = LocalDate.now().getDayOfWeek().toString();
        String sheetName = getSheetName(today);
        System.out.println("Today's Sheet: " + sheetName);

        List<String> keywords = readKeywordsFromExcel(sheetName);
        if (keywords.isEmpty()) {
            System.out.println("No keywords found for today: " + sheetName);
            return;
        }

        List<String[]> results = new ArrayList<>();
        for (String keyword : keywords) {
            driver.get("https://www.google.com");
            WebElement searchBox = driver.findElement(By.xpath("//textarea[@name='q']"));
            searchBox.sendKeys(keyword);
            Thread.sleep(2000);

            List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']//li//span"));
            List<String> suggestionTexts = new ArrayList<>();
            for (WebElement suggestion : suggestions) {
                String text = suggestion.getText();
                if (!text.isEmpty()) {
                    suggestionTexts.add(text);
                }
            }

            if (suggestionTexts.isEmpty()) {
                System.out.println("No suggestions found for keyword: " + keyword);
                continue;
            }

            String longest = suggestionTexts.stream().max((a, b) -> Integer.compare(a.length(), b.length())).orElse(keyword);
            String shortest = suggestionTexts.stream().min((a, b) -> Integer.compare(a.length(), b.length())).orElse(keyword);

            results.add(new String[]{longest, shortest});
        }

        writeResultsToExcel(sheetName, results);
        System.out.println("Excel updated successfully.");
    }

    @AfterSuite
    public void closeBrowser() {
        driver.quit();
    }

    private String getSheetName(String dayOfWeek) {
        switch (dayOfWeek.toUpperCase()) {
            case "SATURDAY": return "Saturday";
            case "SUNDAY": return "Sunday";
            case "MONDAY": return "Monday";
            case "TUESDAY": return "Tuesday";
            case "WEDNESDAY": return "Wednesday";
            case "THURSDAY": return "Thursday";
            case "FRIDAY": return "Friday";
            default: return "Monday";
        }
    }

    private List<String> readKeywordsFromExcel(String sheetName) {
        List<String> keywords = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(excelFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("Error: Sheet '" + sheetName + "' not found!");
                return keywords;
            }

            for (Row row : sheet) {
                if (row.getRowNum() < 2) continue;
                Cell cell = row.getCell(2);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    keywords.add(cell.getStringCellValue().trim());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return keywords;
    }

    private void writeResultsToExcel(String sheetName, List<String[]> results) {
        try (FileInputStream file = new FileInputStream(excelFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("Error: Sheet '" + sheetName + "' not found while writing!");
                return;
            }

            for (int i = 0; i < results.size(); i++) {
                Row row = sheet.getRow(i + 2);
                if (row == null) row = sheet.createRow(i + 2);
                row.createCell(3, CellType.STRING).setCellValue(results.get(i)[0]);
                row.createCell(4, CellType.STRING).setCellValue(results.get(i)[1]);
            }

            try (FileOutputStream outFile = new FileOutputStream(excelFilePath)) {
                workbook.write(outFile);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
