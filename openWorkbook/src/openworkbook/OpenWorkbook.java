package openworkbook;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

public class OpenWorkbook {

    public static void main(String[] args) {
        //Creating a file pointer to Excel file to open
        File file = new File("openWorkbook.xlsx");   //the string should be the location of the file.
        
        try{   //FileNotFounException
            FileInputStream in = new FileInputStream(file);
            //Get the Workbook instance for xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            if(file.exists() && file.isFile()){
                System.out.println("openWorkbook.xlsx opened successfully.");
            }
            else{
                System.out.println("Error in opening the file.");
            }
        }
        catch(Exception e){
            e.printStackTrace();
        }
        
    }
    
}
