package test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class LongestSuggestionSearchTest extends BaseTest {

    private final String EXCEL_FILE_PATH = "src/test/resources/Assignment_Excel.xlsx";

    @Test
    public void performKeywordSearch() throws IOException {
        // Step 1: Read the Excel file
        FileInputStream fis = new FileInputStream(EXCEL_FILE_PATH);
        Workbook workbook = new XSSFWorkbook(fis);

        // Step 2: Determine the current day of the week
        DayOfWeek currentDay = LocalDate.now().getDayOfWeek();
        String currentSheetName = currentDay.name().substring(0, 1).toUpperCase() + currentDay.name().substring(1).toLowerCase();

        Sheet sheet = workbook.getSheet(currentSheetName);
        if (sheet == null) {
            System.out.println("Sheet for the current day does not exist.");
            workbook.close();
            fis.close();
            return;
        }

        // Step 3: Iterate through the rows from 3 to 12 (1-based, so 2 to 11 in 0-based)
        for (int i = 2; i <= 11; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue; // Skip if the row is empty

            // Access column C (index 2) for the keyword
            Cell keywordCell = row.getCell(2);
            if (keywordCell != null) {
                String keyword = keywordCell.getStringCellValue();
                System.out.println(keyword);

                // Step 4: Type the keyword in the search box
                WebElement searchBox = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("q")));
                searchBox.clear();
                searchBox.sendKeys(keyword);

                // Step 5: Extract auto-suggestions
                List<WebElement> suggestions = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//ul[@role='listbox']//li")));
                if (suggestions.isEmpty()) {
                    System.out.println("No suggestions found for keyword: " + keyword);
                    // Write "No suggestions found" to Excel if no suggestions
                    Cell titleCell = row.createCell(3); // D column
                    titleCell.setCellValue("No suggestions found for keyword: " + keyword);
                    continue;
                }

                // Step 6: Find the longest visible suggestion
                WebElement longestOptionElement = suggestions.stream()
                        .filter(WebElement::isDisplayed) // Ensure suggestion is visible
                        .max(Comparator.comparingInt(s -> s.getText().length())) // Find longest suggestion by text length
                        .orElse(null);

                if (longestOptionElement == null) {
                    System.out.println("No visible suggestion found for keyword: " + keyword);
                    continue;
                }

                String longestOptionText = longestOptionElement.getText();
                System.out.println("Longest visible suggestion for keyword: " + keyword + " -> " + longestOptionText);

                // Step 7: Click the longest suggestion and get the page title
                String longestPageTitle = "";
                try {
                    longestOptionElement.click();

                    // Wait until the page title is updated and not empty
                    wait.until((ExpectedCondition<Boolean>) driver -> {
                        String title = driver.getTitle();
                        return title != null && !title.isEmpty();
                    });

                    longestPageTitle = driver.getTitle();
                    System.out.println("Title from longest visible suggestion for keyword: " + keyword + " -> " + longestPageTitle);

                    // Remove " - Google Search" from the title if it exists
                    longestPageTitle = longestPageTitle.replace(" - Google Search", "");

                    // Step 8: Write the page title to column D
                    Cell titleCell = row.createCell(3); // D column
                    titleCell.setCellValue(longestPageTitle);

                    // Navigate back
                    driver.navigate().back();
                } catch (Exception e) {
                    System.out.println("Failed to retrieve title for keyword: " + keyword);
                    // If clicking the suggestion fails, write an error message to Excel
                    Cell titleCell = row.createCell(3); // D column
                    titleCell.setCellValue("Failed to retrieve title for keyword: " + keyword);
                }
            }
        }

        // Step 9: Save the workbook with changes
        FileOutputStream fos = new FileOutputStream(EXCEL_FILE_PATH);
        workbook.write(fos);

        // Close resources
        fos.close();
        workbook.close();
        fis.close();
    }
}