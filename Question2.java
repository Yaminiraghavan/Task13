package Task13;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;



public class Question2 {
    public static void main(String[] args) throws FileNotFoundException {
        File f =new File("C:\\Users\\ELCOT\\IdeaProjects\\JavaTask\\src\\test\\resources");
        Workbook book = new XSSFWorkbook();
        Sheet sh =book.createSheet("Sheet 1");
        System.out.println("Excel File has been created successfully.");
        


    }
}
