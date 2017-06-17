package readsheet;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class ReadSheet {

    public static void main(String[] args) {
        //Create an instance of a File object
        File file = new File("writeSheet.xlsx");
        if(file.exists() && file.isFile());
        else
            return;
        
        try{
            FileInputStream in = new FileInputStream(file);
            //XSSF object for parsing and managing Excel file data
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            
            
            while(rowIterator.hasNext()){
                XSSFRow row = (XSSFRow)rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while(cellIterator.hasNext()){
                    XSSFCell cell = (XSSFCell)cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println();
            }
            
            //Close the stream
            in.close();
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }
    
}
