package Task13;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Question3 {
    public static void main(String[] args) throws IOException {

        File f = new File("C:\\Users\\ELCOT\\IdeaProjects\\JavaTask\\src\\test\\resources\\Sheet 1.xlsx");
        Workbook book = new XSSFWorkbook();
        Sheet sh = book.createSheet("Sheet 1");
        Row row = sh.createRow(0);
        row.createCell(0).setCellValue("Name");
        row.createCell(1).setCellValue("Age");
        row.createCell(2).setCellValue("Email");

        Row row1 = sh.createRow(1);
        row1.createCell(0).setCellValue("John Doe");
        row1.createCell(1).setCellValue("30");
        row1.createCell(2).setCellValue("john@test.com");

        Row row2 = sh.createRow(2);
        row2.createCell(0).setCellValue("Jane Doe");
        row2.createCell(1).setCellValue("28");
        row2.createCell(2).setCellValue("john@test.com");

        Row row3 = sh.createRow(3);
        row3.createCell(0).setCellValue("Bob Smith");
        row3.createCell(1).setCellValue("35");
        row3.createCell(2).setCellValue("jacky@example.com");

        Row row4 = sh.createRow(4);
        row4.createCell(0).setCellValue("Swapnil");
        row4.createCell(1).setCellValue("37");
        row4.createCell(2).setCellValue("swapnil@example.com");

        FileOutputStream out = new FileOutputStream(f);
        book.write(out);
        book.close();
        System.out.println("Excel file has been generated successfully.");



        /*Map<String,Object[] > mp =new TreeMap();
        mp.put( "1", new Object[] {
                "Name","Age", "Email" });
        mp.put( "2", new Object[] {
                "John Doe", 30, "john@test.com" });

       mp.put( "3", new Object[] {
                "Bob Smith", 35, "jacky@example.com" });

       mp.put( "4", new Object[] {
               "Swapnil", 37, "swapnil@example.com" });
        mp.put( "5", new Object[] {
                "Jane Doe", 28, "john@test.com" });

        Set<String> S=mp.keySet();
        int rowid = 0;
        for (String key: S) {
            Row row = sh.createRow(rowid++);
            Object[] Arr = mp.get(key);
        int cellid=0;
        for (Object ob:Arr){
            Cell cell = row.createCell(cellid++);
            cell.setCellValue((String)ob);
        }
        }
        FileOutputStream out = new FileOutputStream(
                new File("C:\\Users\\ELCOT\\IdeaProjects\\JavaTask\\src\\test\\resources\\Sheet 1.xlsx"));

        book.write(out);
        out.close();
        System.out.println("Sheet1.xlsx written successfully");*/


    }
}
