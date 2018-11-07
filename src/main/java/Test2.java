import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Test2 {
    public static void main(String[] args) throws IOException {
        File file = new File("C:\\Users\\Administrator\\Desktop\\新建文件夹\\2017题库\\test.xlsx");
        FileInputStream in = null;

        try {
            in = new FileInputStream(file); // 文件流
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Workbook workbook = new XSSFWorkbook(in);
        Sheet sheet = workbook.getSheetAt(0);

        int allRowNum = sheet.getLastRowNum();
        Row row = sheet.getRow(1);
//        System.out.println(row.getCell(0));
        int end = row.getLastCellNum();
        for (int i = 0; i<end;i++){
            Cell cell = row.getCell(i);
            Object object = getValue(cell);
            System.out.println(object);
        }

    }
    private static Object getValue(Cell cell) {
        Object obj = null;
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case ERROR:
                obj = cell.getErrorCellValue();
                break;
            case NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case STRING:
                obj = cell.getStringCellValue();

                break;
            default:
                break;
        }
        return obj;
    }
}
