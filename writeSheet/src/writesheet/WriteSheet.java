package writesheet;

import java.io.*;   //adds File and FileOutputStream
import org.apache.poi.xssf.usermodel.*;     //adds XSSFWorkbook,XSSFSheet,XSSFRow and XSSFCell
import java.util.*;     //adds Set,Map,TreeMap

public class WriteSheet {

    public static void main(String[] args) {
        //Creating an instance for the workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Creating an instance for the sheet 
        XSSFSheet sheet = workbook.createSheet("Grocery");
        
        
        //Adding data that should be in the Excel sheet
        Map<String,Object[]> groceryItem = new TreeMap<String,Object[]>();
        groceryItem.put("1", new Object[] {"Item", "Price(Rs./Kg)", "Quantity(Kg)"});
        groceryItem.put("2", new Object[] {"Apple", 60, 2});
        groceryItem.put("3", new Object[] {"Moong Daal", 75, 1});
        groceryItem.put("4", new Object[] {"Biscuits", 40, 2});
        groceryItem.put("5", new Object[] {"Orange", 50, 2});
        groceryItem.put("6", new Object[] {"Soap", 10, 0.25});
        groceryItem.put("7", new Object[] {"Chana", 40, 2});
        
        
        //Iterating over data and writing to sheet
        Set<String> keyset = groceryItem.keySet();  
        int rownum = 0;     //rows are 0 indexed
        for(String key : keyset){
            //Creating row instance
            XSSFRow row = sheet.createRow(rownum);
            rownum++;   //increment to next row
            Object [] objectArr = groceryItem.get(key);
            int cellnum = 0;
            for(Object obj : objectArr){
                XSSFCell cell = row.createCell(cellnum);
                cellnum++;
                if(obj instanceof String){
                    cell.setCellValue((String)obj);
                }
                else if(obj instanceof Integer){
                    cell.setCellValue((Integer)obj);
                }
                else if(obj instanceof Double){
                    cell.setCellValue((Double)obj);
                }
            }
        }
        

        try{
            FileOutputStream out = new FileOutputStream( new File("writeSheet.xlsx") );
            workbook.write(out);
            out.close();
            System.out.println("writeSheet.xlsx created successfully. Data has been written.");
        }
        catch(Exception e){   //Handles exception like FileNotFoundException
            e.printStackTrace();
        }
    }
    
}
