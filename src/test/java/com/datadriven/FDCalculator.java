package com.datadriven;

import java.time.Duration;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class FDCalculator {

    public static void main(String[] args) throws Exception {

        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html");
        driver.manage().window().maximize();

        String filePath = System.getProperty("user.dir") + "\\testdata\\FDData.xlsx";

        int rows = ExcelUtils.getRowCount(filePath, "Sheet1");

        for (int i = 1; i <= rows; i++) {

            // Read data from Excel
            String principal = ExcelUtils.getCellData(filePath, "Sheet1", i, 0);
            String rate = ExcelUtils.getCellData(filePath, "Sheet1", i, 1);
            String period1 = ExcelUtils.getCellData(filePath, "Sheet1", i, 2);
            String period2 = ExcelUtils.getCellData(filePath, "Sheet1", i, 3);
            String frequency = ExcelUtils.getCellData(filePath, "Sheet1", i, 4);
            String expectedMaturityAmount = ExcelUtils.getCellData(filePath, "Sheet1", i, 5);

            try {
                // Clear and enter Principal
                if (principal.contains(".")) {
                    principal = principal.substring(0, principal.indexOf("."));
                }
                WebElement principalElement = driver.findElement(By.id("principal"));
                principalElement.clear();
                principalElement.sendKeys(principal);

                // Enter Interest
                WebElement intr = driver.findElement(By.id("interest"));
                intr.clear();
                intr.sendKeys(rate);

                // Clear and enter Tenure
                if (period1.contains(".")) {
                    period1 = period1.substring(0, period1.indexOf("."));
                }
                WebElement tenure = driver.findElement(By.id("tenure"));
                tenure.clear();
                tenure.sendKeys(period1);

                // Handle alert if any
                try {
                    WebDriverWait shortWait = new WebDriverWait(driver, Duration.ofSeconds(2));
                    shortWait.until(ExpectedConditions.alertIsPresent());
                    Alert alert = driver.switchTo().alert();
                    System.out.println("Alert detected: " + alert.getText());
                    alert.accept();
                    ExcelUtils.setCellData(filePath, "Sheet1", i, 7, "Failed - Alert");
                    ExcelUtils.fillRedColor(filePath, "Sheet1", i, 7);
                    continue;
                } catch (TimeoutException | NoAlertPresentException e) {
                    // No alert present, continue
                }

                // Select dropdowns
                Select periodDropdown = new Select(driver.findElement(By.id("tenurePeriod")));
                periodDropdown.selectByVisibleText(period2);

                Select frequencyDropdown = new Select(driver.findElement(By.id("frequency")));
                frequencyDropdown.selectByVisibleText(frequency);

                // Wait for overlay to disappear
                try {
                    wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector(".wzrk-overlay")));
                } catch (Exception e) {
                    System.out.println("Overlay not found or timed out.");
                }

                // Click Calculate button using JS
                WebElement calculateBtn = driver.findElement(By.xpath("//div[@class='cal_div']//a[1]"));
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", calculateBtn);

                // Wait for result
                Thread.sleep(1000);
                String actualMaturityAmount = driver.findElement(By.xpath("//span[@id='resp_matval']//strong")).getText();

                // Validation
                if (Double.parseDouble(actualMaturityAmount) == Double.parseDouble(expectedMaturityAmount)) {
                    System.out.println("Test case " + i + " passed. Expected: " + expectedMaturityAmount + ", Actual: " + actualMaturityAmount);
                    ExcelUtils.setCellData(filePath, "Sheet1", i, 7, "Passed");
                    ExcelUtils.fillGreenColor(filePath, "Sheet1", i, 7);
                } else {
                    System.out.println("Test case " + i + " failed. Expected: " + expectedMaturityAmount + ", Actual: " + actualMaturityAmount);
                    ExcelUtils.setCellData(filePath, "Sheet1", i, 7, "Failed");
                    ExcelUtils.fillRedColor(filePath, "Sheet1", i, 7);
                }

                // Reset form
                Thread.sleep(2000);
                WebElement resetBtn = driver.findElement(By.xpath("//img[@class='PL5']"));
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", resetBtn);

            } catch (Exception ex) {
                System.out.println("Exception in test case " + i + ": " + ex.getMessage());
                ExcelUtils.setCellData(filePath, "Sheet1", i, 7, "Error - Exception");
                ExcelUtils.fillRedColor(filePath, "Sheet1", i, 7);
            }
        }

        driver.quit();
    }
}
