import java.io.FileOutputStream;
import java.util.Hashtable;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Assertions;

public class CreateExcel {
    public static void CreateSimple(String[] args) throws Throwable {

        // keep 10 rows in memory, exceeding rows will be flushed to disk
        SXSSFWorkbook wb = new SXSSFWorkbook(10);

        Sheet sh = wb.createSheet();

        Hashtable ht = new Hashtable();
        ht.put(0, "Paris");
        ht.put(1, "Marseille");
        ht.put(2, "Lille");
        ht.put(3, "Bordeaux");


        for(int rownum = 0; rownum < 100; rownum++){
            Row row = sh.createRow(rownum);
            for(int cellnum = 0; cellnum < ht.size(); cellnum++){
                Cell cell = row.createCell(cellnum);
                //String address = new CellReference(cell).formatAsString();
                String address = ht.get(cellnum).toString() + rownum;
                cell.setCellValue(address);
            }
        }


        // Rows with rownum < 90 are flushed and not accessible
        for(int rownum = 0; rownum < 90; rownum++){
            Assertions.assertNull(sh.getRow(rownum));
        }


        // ther last 10 rows are still in memory
        for(int rownum = 90; rownum < 100; rownum++){
            Assertions.assertNotNull(sh.getRow(rownum));
        }

        String name = "C:\\Users\\alexe\\sncf\\sncf_first_task\\fichiers_excel\\task_1_create_simple\\test.xslx";
        FileOutputStream out = new FileOutputStream(name);

        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }
}
