package cellTypes;

import java.io.*;
import java.util.Date;
import org.apache.poi.xssf.usermodel.*;

public class CellTypes {

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("Type of Cell");
        row.createCell(1).setCellValue("Cell value");
        row = sheet.createRow((short) 1);
        row.createCell(0).setCellValue("Set cell type BLANK");  //XSSFCell.CELL_TYPE_BLANK
        row.createCell(1);
        row = sheet.createRow((short) 2);
        row.createCell(0).setCellValue("Set cell type BOOLEAN");    //XSSFCell.CELL_TYPE_BOOLEAN
        row.createCell(1).setCellValue(false);
        row = sheet.createRow((short) 5);
        row.createCell(0).setCellValue("Set cell type ERROR");  //XSSFCell.CELL_TYPE_ERROR
        row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR );
        row = sheet.createRow((short) 6);
        row.createCell(0).setCellValue("Set cell type date");   
        row.createCell(1).setCellValue(new Date());
        Date d = new Date();
        System.out.println(d);
        row = sheet.createRow((short) 7);
        row.createCell(0).setCellValue("Set cell type numeric" );   //XSSFCell.CELL_TYPE_NUMERIC
        row.createCell(1).setCellValue(20 );
        row = sheet.createRow((short) 8);
        row.createCell(0).setCellValue("Set cell type string");   //XSSFCell.CELL_TYPE_STRING
        row.createCell(1).setCellValue("A String");
        
        try{
            FileOutputStream out = new FileOutputStream( new File("cellTypes.xlsx") );
            workbook.write(out);
            out.close();
            System.out.println("cellTypes.xlsx created successfully.");
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }
    
}
