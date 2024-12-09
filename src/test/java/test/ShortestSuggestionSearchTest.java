package test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

public class ShortestSuggestionSearchTest extends BaseTest {

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
                    Cell titleCell = row.createCell(4); // E column
                    titleCell.setCellValue("No suggestions found for keyword: " + keyword);
                    continue;
                }

                // Step 6: Find the shortest visible suggestion
                WebElement shortestOption = null;
                int shortestLength = Integer.MAX_VALUE;

                for (WebElement suggestion : suggestions) {
                    if (suggestion.isDisplayed()) {
                        String suggestionText = suggestion.getText();
                        if (suggestionText.length() < shortestLength) {
                            shortestLength = suggestionText.length();
                            shortestOption = suggestion;
                        }
                    }
                }

                if (shortestOption == null) {
                    System.out.println("No valid visible suggestions found for keyword: " + keyword);
                    // Write message to Excel if no valid suggestion found
                    Cell titleCell = row.createCell(4); // E column
                    titleCell.setCellValue("No valid suggestions found for keyword: " + keyword);
                    continue;
                }

                // Step 7: Select the shortest visible option
                shortestOption.click();

                // Wait until the page title is updated and not empty
                wait.until(driver -> {
                    String title = driver.getTitle();
                    return title != null && !title.isEmpty();
                });

                // Get the page title
                String pageTitle = driver.getTitle();
                System.out.println("Page title for shortest option for keyword: " + keyword + " -> " + pageTitle);

                // Remove " - Google Search" from the title if it exists
                pageTitle = pageTitle.replace(" - Google Search", "");

                // Step 8: Write the page title to column E
                Cell titleCell = row.createCell(4); // E column
                titleCell.setCellValue(pageTitle);

                // Step 9: Navigate back to the search page
                driver.navigate().back();

                // Step 10: Clear the search box for the next keyword
                WebElement searchBoxAfterBack = wait.until(ExpectedConditions.presenceOfElementLocated(By.name("q")));
                searchBoxAfterBack.clear();
            }
        }

        // Step 11: Save the workbook with changes
        FileOutputStream fos = new FileOutputStream(EXCEL_FILE_PATH);
        workbook.write(fos);

        // Close resources
        fos.close();
        workbook.close();
        fis.close();
    }
}