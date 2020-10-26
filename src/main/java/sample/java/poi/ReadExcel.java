package sample.java.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Hello world!
 *
 */
public class ReadExcel {

    public static final String FILE_NAME = "test.xlsx";
    public static final String SHEET_NAME = "Sheet1";

    public static final int INIT_ROW = 0;
    public static final int INIT_COLUMN = 0;

    public static void main(String[] args)
    throws EncryptedDocumentException, InvalidFormatException, IOException {

        FileInputStream in
            = new FileInputStream(
                Paths.get(System.getProperty("user.dir"), "target", "classes", FILE_NAME).toString());;
        Workbook book = WorkbookFactory.create(in);
        Sheet sheet = book.getSheet(SHEET_NAME);

        int rowIndex = INIT_ROW;
        int columnIndex = INIT_COLUMN;

        printCellValue(sheet, rowIndex, columnIndex);

        rowIndex++;
        printCellValue(sheet, rowIndex, columnIndex);

        book.close();
        in.close();
    }

    private static void printCellValue(Sheet sheet, int rowIndex , int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        System.out.println("Cell:" + getCell(rowIndex, columnIndex) + ", value:" + cell.getStringCellValue());
    }

    private static String getCell(int rowIndex, int columnIndex){
        return getColumn(columnIndex) + getRow(rowIndex);
    }

    private static String getColumn(int index) {
        return Character.toString((char)('A' + index));
    }

    private static String getRow(int index){
        return Integer.toString(index + 1);
    }
}
