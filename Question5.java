package Task13;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Question5 {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "C:\\Users\\ELCOT\\IdeaProjects\\JavaTask\\src\\test\\resources\\Sheet 1.xlsx";
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(getCellValue(cell) + "\t");
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "Unknown Cell Type";
        }
        /* FormulaEvaluator formulaEvaluator=book.getCreationHelper().createFormulaEvaluator();
       for(Row row: sh)
        {
            for(Cell cell: row)
            {
                switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                {
                    case NUMERIC:

                        System.out.print(cell.getNumericCellValue()+ "\t\t");
                        break;
                    case STRING:

                        System.out.print(cell.getStringCellValue()+ "\t\t");
                        break;
                }
            }
            System.out.println();

        }*/
    }
}



