package poi;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;


public class poi {
    public static void main(String[] args) {
        try {
            InputStream is = new FileInputStream(new File("/Users/benashby/Documents/JavaTraining/Maven1/src/main/resources/lims.xlsx"));
            Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(100)
                    .bufferSize(4096)
                    .open(is);
            {
                for (Sheet sheet : workbook) {
                    System.out.println(sheet.getSheetName());
                    for (Row r : sheet) {
                        for (Cell c : r) {
                            System.out.println(c.getStringCellValue());
                            System.out.println(c.getColumnIndex());
                        }
                    }
                }
            }
        } catch (FileNotFoundException ex) {
            // insert code to run when exception occurs
        }


    }
}
