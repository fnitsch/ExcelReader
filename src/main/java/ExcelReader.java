import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    private Workbook workbook;
    private Sheet sheet;

    private FileInputStream fileInputStream;

    public void openFile(String filePath) throws EncryptedDocumentException, IOException {

        File file = new File(
                getClass().getClassLoader().getResource(filePath).getFile()
        );

        fileInputStream = new FileInputStream(file);

        workbook = WorkbookFactory.create(fileInputStream);

        sheet = workbook.getSheet("Tabelle1");

        int lastRowNum = sheet.getLastRowNum();

        for (int i = 1; i <= lastRowNum; i++) {

            System.out.println("ID: " + sheet.getRow(i).getCell(0) + ", Vorname: " + sheet.getRow(i).getCell(1) + ", Nachname: " + sheet.getRow(i).getCell(2));
        }

        fileInputStream.close();
        workbook.close();
    }

}
