package LIS_Application_Testing.LIS;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

import java.io.FileInputStream;
import java.io.IOException;

public class Login {

    WebDriver driver;

    @Test(priority = 0)
    public void launchbrowser() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver(); //Invoke Chrome Browser
        driver.manage().window().maximize();
    }

    @Test(priority = 1)
    public void loginDatadriven() throws IOException {
        driver.get("http://127.0.0.1:8000/login");

        //Getting Data from Excel Sheet
        FileInputStream file = new FileInputStream("C:\\Users\\admin\\Documents\\MLCApplication\\logindatadriventesting.xlsx");

        //Getting the Workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        //Get the first sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);

        //Get the number of rows
        int noOfRows = sheet.getLastRowNum(); //Get all rows from sheet or Row count

        System.out.println("No of Records in the Excel Sheet: " + noOfRows);

        for (int row = 1; row <= noOfRows; row++) {

            XSSFRow current_row = sheet.getRow(row);  //First Row

            String email = getCellValueAsString(current_row.getCell(0));
            String password = getCellValueAsString(current_row.getCell(1));

            //Entering Login Information
            driver.findElement(By.xpath("//input[@id='email']")).sendKeys(email);
            driver.findElement(By.xpath("//input[@id='password']")).sendKeys(password);
            driver.findElement(By.xpath("//button[contains(text(),'Login')]")).click();

            if (driver.getPageSource().contains("Dashboard")) {
                System.out.println("Login successful for record " + row);
            } else {
                System.out.println("Login failed for record " + row);
            }
        }
        System.out.println("Data Driven Test Completed");
        driver.close();
        driver.quit();
        file.close(); //Close Excel File
    }

    private String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        } else {
            return "";
        }
    }

    public static void main(String[] args) throws IOException {
        Login l = new Login();
        l.launchbrowser();
        l.loginDatadriven();
    }
}
