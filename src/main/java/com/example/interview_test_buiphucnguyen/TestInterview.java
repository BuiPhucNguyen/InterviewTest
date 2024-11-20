package com.example.interview_test_buiphucnguyen;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AddSheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.auth.Credentials;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.ServiceAccountCredentials;


public class TestInterview {    
	private static final String APPLICATION_NAME = "Google Sheets API Service Account Example";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String SPREADSHEET_ID = "1HEx7xSmJFzQqidr_4FzFQVE-XJRG9fjo1FZTGkEDl4Y"; // ID của bảng tính Google Sheets
    private static final String CREDENTIALS_FILE_PATH = "D:\\Wordspaces\\SeleniumJava_Eclipse2022\\interview-test-buiphucnguyen\\json\\interview-test-442314-8eba86df8670.json"; // Đường dẫn đến tệp JSON của Service Account
    
    private static Sheets sheetsService;
    
    private static Sheets getSheetsService() throws Exception {
        System.out.println("Starting Sheets service initialization...");

        // Kiểm tra tệp credentials
        File credentialsFile = new File(CREDENTIALS_FILE_PATH);
        if (!credentialsFile.exists()) {
            throw new IOException("Credential file not found: " + CREDENTIALS_FILE_PATH);
        }
        System.out.println("Credentials file found at: " + CREDENTIALS_FILE_PATH);

        // Xác thực với Service Account
        try (FileInputStream serviceAccountStream = new FileInputStream(credentialsFile)) {
            Credentials credentials = ServiceAccountCredentials.fromStream(serviceAccountStream)
                    .createScoped(Collections.singletonList("https://www.googleapis.com/auth/spreadsheets"));

            System.out.println("Credentials loaded successfully.");

            // Tạo đối tượng Sheets service
            Sheets sheetsService = new Sheets.Builder(
                    GoogleNetHttpTransport.newTrustedTransport(),
                    JSON_FACTORY,
                    new HttpCredentialsAdapter(credentials))
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            System.out.println("Sheets service initialized successfully.");
            return sheetsService;
        } catch (Exception e) {
            e.printStackTrace();
            throw new IllegalStateException("Failed to initialize Sheets service: " + e.getMessage());
        }
    }

    
    private static void createSheet(String sheetName) throws IOException {
        if (sheetsService == null) {
            throw new IllegalStateException("Sheets service has not been initialized.");
        }

        AddSheetRequest addSheetRequest = new AddSheetRequest()
                .setProperties(new SheetProperties().setTitle(sheetName));

        BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(Arrays.asList(new Request().setAddSheet(addSheetRequest)));

        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, batchUpdateRequest).execute();
        System.out.println("Sheet '" + sheetName + "' created successfully.");
    }
    
    private static void updateSheet(Sheets sheetsService, String sheetName, List<List<Object>> values) throws IOException {
        ValueRange body = new ValueRange().setValues(values);
        String range = sheetName + "!A1";
        UpdateValuesResponse result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, range, body)
                .setValueInputOption("RAW")
                .execute();

        System.out.println(result.getUpdatedCells() + " cells updated.");
    }
    
    public static List<List<Object>> convertToSheetData(List<Player> players) {
        List<List<Object>> sheetData = new ArrayList<>();

        // Thêm tiêu đề cột (header) vào danh sách dữ liệu
        List<Object> header = new ArrayList<>();
        header.add("#");
        header.add("Name");
        header.add("Title");
        header.add("Fed");
        header.add("Rating");
        header.add("G");
        header.add("B-Year");
        sheetData.add(header);

        // Thêm dữ liệu của từng đối tượng Person
        for (Player person : players) {
            List<Object> row = new ArrayList<>();
            row.add(person.getNumber());
            row.add(person.getName());
            row.add(person.getTitle());
            row.add(person.getFed());
            row.add(person.getRating());
            row.add(person.getG());
            row.add(person.getByear());
            sheetData.add(row);
        }

        return sheetData;
    }
	
	public static void main(String[] args) throws Exception {
		
		 System.setProperty("webdriver.chrome.driver", "D:\\SeleniumJava\\WebDriver\\chromedriver.exe");
		  
		  WebDriver driver = new ChromeDriver();
		    
	      driver.manage().window().maximize();
	      driver.manage().timeouts().pageLoadTimeout(20, TimeUnit.SECONDS);
	      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	      
	      driver.navigate().to("https://ratings.fide.com/topfed.phtml?ina=1&country=AUS");
	      Thread.sleep(2000);
	      
	      List<WebElement> totalRows = driver.findElements(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr"));
	      System.out.println("Số dòng tìm thấy: " + totalRows.size());
	      
	      List<WebElement> totalColumns = driver.findElements(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr[1]/td"));
	      System.out.println("Số cột tìm thấy: " + totalColumns.size());
	      
	      List<Player> players = new ArrayList<>();
	      
	      //Lấy 100 dòng đầu 
	      //int i = 2 để bỏ dòng tiêu đề cột
	      for (int i = 2; i <= totalRows.size(); i++) {
	    	  
				WebElement number = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[1]/font"));
				
				JavascriptExecutor js = (JavascriptExecutor) driver;
	            js.executeScript("arguments[0].scrollIntoView(true);", number);
				
				WebElement name = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[2]/a"));
				WebElement title = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[3]"));
				WebElement fed = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[4]"));
				WebElement rating = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[5]"));
				WebElement g = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[6]"));
				WebElement byear = driver.findElement(By.xpath(" //tbody/tr/td[@valign='top']/table[2]//tbody/tr["+i+"]/td[7]"));
				
				Player player = new Player();
				player.setNumber(Integer.parseInt(number.getText().trim()));
				player.setName(name.getText().trim());
				player.setTitle(title.getText().trim());
				player.setFed(fed.getText().trim());
				player.setRating(Integer.parseInt(rating.getText().trim()));
				player.setG(Integer.parseInt(g.getText().trim()));
				player.setByear(Integer.parseInt(byear.getText().trim()));
				
				players.add(player);
			
	      }
	      
	      System.out.println("Danh sách players:");
	      for (Player player : players) {
	    	  System.out.println(player.toString());
	      }
	      
	      Thread.sleep(2000);
	      driver.quit();
	      
	      // Xác thực và tạo Sheets service
	      sheetsService = getSheetsService();
	      
	      List<List<Object>> sheetData = convertToSheetData(players);
	      
	      // Tạo sheet mới
	      String sheetName = "NewSheet"; // Tên sheet mới bạn muốn tạo
	      createSheet(sheetName);
	      
	      // Ghi vào sheet mới
	      // Cập nhật dữ liệu vào sheet
	      updateSheet(sheetsService, sheetName, sheetData);      
	}
}
