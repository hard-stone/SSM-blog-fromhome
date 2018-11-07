import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

//https://blog.csdn.net/wwd0501/article/details/78780646
public class Test {

    public static void main(String[] args) throws IOException {

        String EXCELDIR="C:\\Users\\Administrator\\Desktop\\新建文件夹\\2017题库";
        File f = new File(EXCELDIR);
        for (File temp : f.listFiles()) {
            if (temp.isFile()) {
                String name = temp.getName();
                String wordname = name.substring(0,name.lastIndexOf("."))+".docx";
                String excelpath = EXCELDIR+"\\"+name;
                System.out.println(wordname);
                System.out.println(excelpath);
                File excelFile = new File(excelpath); // 创建文件对象
                writeword(readexcel(excelFile),wordname);
            }
        }
//
//        File file = new File("C:\\Users\\Administrator\\Desktop\\新建文件夹\\2017题库\\安检信息管理系统-题库.xlsx");
//        List list = readexcel(file);
//        writeword(list,"test.docx");
    }

    private static  List<List> readexcel(File file) throws IOException {
        List<List> list = new ArrayList<List>();
        FileInputStream in = null;

        try {
            in = new FileInputStream(file); // 文件流
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Workbook workbook = new XSSFWorkbook(in);
//        int sheetCount = workbook.getNumberOfSheets();
//        System.out.println("sheetCount :" + sheetCount);
        Sheet sheet = workbook.getSheetAt(0);
        int allRowNum = sheet.getLastRowNum();
//        int allColumns = sheet.get;
        System.out.println("总行数: " + sheet.getLastRowNum());
        int count = 0;
        for (Row row : sheet) {
//            int columnTotalnum = row.getPhysicalNumberOfCells();
            int end = row.getLastCellNum();
//            System.out.println(columnTotalnum);
//            System.out.println(row.getLastCellNum());
//            System.out.println(row.getCell(0).toString());
            List columnlist = new ArrayList();
            for (int i = 0; i < end; i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    columnlist.add("空");
//                    System.out.println("空"+"\t");
                    continue;
                }
                Object object = getValue(cell);
                columnlist.add(object);
//                System.out.println(object);
            }
            list.add(columnlist);
            System.out.println("\t");
        }
        for (int i = 0; i < list.size(); i++) {
            List temp = list.get(i);
            for (int j = 0; j < temp.size(); j++) {
                System.out.print(temp.get(j)+",");
            }
            System.out.println("\t");
        }
        return  list;
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

    private static void writeword(List<List> lists,String filename) {
        XWPFDocument document = new XWPFDocument();
        OutputStream stream = null;
        BufferedOutputStream bufferedOutputStream = null;
        String WORDDIR = "C:\\Users\\Administrator\\Desktop\\新建文件夹\\word\\";
        filename = WORDDIR+filename;
        String[] head = new String[7];
        try {
            stream = new FileOutputStream(new File(filename));
            bufferedOutputStream = new BufferedOutputStream(stream, 1024);
            XWPFParagraph p1 = document.createParagraph();
            p1.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun r1 = p1.createRun();
            r1.setTextPosition(10);
            for (int i = 0; i < lists.get(0).size(); i++) {
                head[i] = String.valueOf(lists.get(0).get(i));
                System.out.println(head[i]);
            }

//            String temp = head[1];
            XWPFParagraph p2 = null;
            XWPFRun r2 = null;
            for (int i = 1; i < lists.size(); i++) {
                ArrayList arr = (ArrayList) lists.get(i);
//                p2 = document.createParagraph();
//                r2 = p2.createRun();

//                if ("空".equals(arr.get(0).toString())) {
//                    r2.setText(arr.get(4).toString());
////                    r2.addCarriageReturn();
//                    r2.setFontSize(15);
//                    continue;
//                }
//                Double num = Double.valueOf(arr.get(0).toString());
//                int a = (int) Math.ceil(num);
//                if ("单选".equals(arr.get(1).toString())) {
//                    r2.setText(a + "、" + arr.get(1) + "题  ");
//                    r2.addCarriageReturn();
//                    r2.setText(arr.get(3) + "");
//                    r2.addCarriageReturn();
//                    r2.setText("答案: " + arr.get(5));
//                    r2.addCarriageReturn();
//                    r2.setText(arr.get(4) + "");
//                    r2.setFontSize(15);
//                    continue;
//                }


//                if (!arr.get(1).equals(temp)){
//                    XWPFParagraph p3 = document.createParagraph();;
//                    XWPFRun r3 = p2.createRun();
//                    temp = (String) arr.get(1);
//                    System.out.println(temp);
//                    r3.setText(temp);
//                    r3.addCarriageReturn();
//                    r3.setFontSize(20);
//                }
                Double num = Double.valueOf(arr.get(0).toString());
                int a = (int) Math.ceil(num);
                p2 = document.createParagraph();
                r2 = p2.createRun();

                r2.setText(a+ "、"+ arr.get(3));
                r2.addCarriageReturn();
                r2.setText("---答案: " + arr.get(5));
                r2.addCarriageReturn();
                r2.setText((String) arr.get(4));
//                r2.addCarriageReturn();
//                r2.setText("知识点："+(String) arr.get(6));
                r2.setFontSize(10);
//                r2.addCarriageReturn();
            }
            document.write(stream);
            stream.close();
            bufferedOutputStream.close();
        } catch (Exception ex) {
            System.out.println(ex.toString()+"HERE");
        }
    }
}
