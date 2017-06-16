package createworkbook;

import java.io.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;

public class CreateWorkbook{
    public static void main(String args[]) {  //Because FileNotFoundException(File),
        //IOException(write,close) would occur
        System.out.println("***********Demo Create Workbook***********");

        //Create Blank Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        try {
             //Create File System using specific name
            FileOutputStream out = new FileOutputStream( new File("createWorkbook.xlsx") );

            //Write operation to worbook using file output object
            workbook.write(out);
            out.close();
            System.out.println("createWorkbook.xlsx has been successfully created");
        }
       catch(Exception e){
           e.printStackTrace();
       }
    }
}