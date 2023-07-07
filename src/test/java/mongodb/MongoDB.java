package mongodb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoClient;
import com.mongodb.client.MongoClients;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.MongoIterable;

public class MongoDB {

	public static void main(String[] args) throws Exception {

		String Url = "mongodb+srv://adminXealei:hNntCLqUSkTxbJel@xealei-qa.1of90.mongodb.net"
				+ "/xealeiqa?retryWrites=true&w=majority";

		MongoClient client = MongoClients.create(Url);

		MongoDatabase database = client.getDatabase("SampleTesting");

		MongoIterable<Document> lstDB = client.listDatabases();
		for (Document x : lstDB) {
			System.out.println(x);
		}

		MongoCollection<Document> collection = database.getCollection("Test_01");

		Document doc1 = new Document("Name", "Divya").append("Age", "24").append("PhoneNum", "1234567890")
				.append("Address", "America");
		Document doc2 = new Document("Name", "Nila").append("Age", "27").append("PhoneNum", "0987654321")
				.append("Address", "Africa");
		List<Document> li = new ArrayList<Document>();
		li.add(doc1);
		li.add(doc2);

		collection.insertMany(li);

		System.out.println("Successfully Inserted");

		// Create a new workbook
		File file = new File("F:\\Xealei-POC\\src\\test\\resources\\Excel\\MongoDB.xlsx");
	
		FileInputStream stream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Data");

		// Fetch MongoDB documents
		FindIterable<Document> documents = collection.find();
        int rowNumber=0;

		// Add headers to the Excel file
        Row headerRow = sheet.createRow(rowNumber++);
        int columnNumber = 0;
        for (String key : documents.first().keySet()) {
            Cell cell = headerRow.createCell(columnNumber++);
            cell.setCellValue(key);
        }
		
// Add data to the Excel file
        for (Document document : documents) {
            Row dataRow = sheet.createRow(rowNumber++);
            columnNumber = 0;
            for (String key : document.keySet()) {
                Cell cell = dataRow.createCell(columnNumber++);
                Object value = document.get(key);
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                } else if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else {
                    cell.setCellValue(value.toString());
                }
            }
        }

// Auto-size columns
        for (int i = 0; i < documents.first().keySet().size(); i++) {
            sheet.autoSizeColumn(i);
        }

// Save the workbook as an Excel file
        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Excel file created successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }

// Close the MongoDB connection
        client.close();
}

}
