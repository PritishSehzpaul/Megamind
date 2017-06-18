package cellstyles;

import java.io.*;
import org.apache.poi.ss.usermodel.*;       //for IndexedColors class
import org.apache.poi.ss.util.CellRangeAddress;     //for CellRangeAddress class
import org.apache.poi.xssf.usermodel.*;

public class CellStyles {

    public static void main(String[] args) {
        XSSFWorkbook workbook =  new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Cell Styles");
        XSSFRow row= sheet.createRow(0);
       
        /**********MERGING CELL**********/
        row.setHeight((short)500);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("CELL MERGING");
        sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 2));
        
        /**********CELL ALIGNMENT**********/
        row = sheet.createRow(4);
        row.setHeight((short)500);
        cell = row.createCell(0);
        //Top Left Alignment
        XSSFCellStyle style1 = workbook.createCellStyle();
        style1.setAlignment(HorizontalAlignment.LEFT);
        style1.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellValue("Cell Alignment: Top Left");
        cell.setCellStyle(style1);
        
        row = sheet.createRow(5);
        row.setHeight((short)500);
        cell = row.createCell(0);
        //Center Alignment
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellValue("Center Aligned");
        cell.setCellStyle(style2);
        
        row = sheet.createRow(6);
        row.setHeight((short)500);
        cell = row.createCell(1);
        //Bottom Right Alignment
        XSSFCellStyle style3 = workbook.createCellStyle();
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellValue("Bottom Right Alignment");
        cell.setCellStyle(style3);
        
        row = sheet.createRow(8);
        row.setHeight((short)800);
        cell = row.createCell(2);
        //Justified Alignment
        XSSFCellStyle style4 = workbook.createCellStyle();
        style4.setAlignment(HorizontalAlignment.JUSTIFY);
        style4.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        cell.setCellValue("This text has been justified");     //To see the effects of justification enter a lign that wraps inside the cell like "I love Maths,Physics,Hope,Philosphy,Happiness,Research,Finding the truth"
        cell.setCellStyle(style4);
        
        /**********CELL BORDER**********/
        row = sheet.createRow(10);
        row.setHeight((short)600);
        cell = row.createCell(1);
        cell.setCellValue("Bordered Cell");
        
        XSSFCellStyle style5 = workbook.createCellStyle();
        style5.setBorderBottom(BorderStyle.THICK);
        style5.setBottomBorderColor(IndexedColors.RED.getIndex());      //use getIndex() or index. Both give same result
        style5.setBorderLeft(BorderStyle.DASHED);
        style5.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        style5.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style5.setTopBorderColor(IndexedColors.BRIGHT_GREEN.index);
        style5.setBorderRight(BorderStyle.DOUBLE);
        style5.setRightBorderColor(IndexedColors.MAROON.index);
        cell.setCellStyle(style5);
        
        /**********FILL COLORS**********/
        //background colors
        row = sheet.createRow(11);
        cell = row.createCell(0);
        row.setHeight((short)800);
        
        XSSFCellStyle style6 = workbook.createCellStyle();
        style6.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        style6.setFillPattern(FillPatternType.LESS_DOTS);
        style6.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style6);
        sheet.setColumnWidth(1,5000);       //sets the thickness or width fo the specified column
        cell.setCellValue("FILL BACKGROUND/FILL PATTERN");
        
        //foreground colors
        row = sheet.createRow(12);
        row.setHeight((short)500);
        cell = row.createCell(1);
        
        XSSFCellStyle style7 = workbook.createCellStyle();
        style7.setFillForegroundColor(IndexedColors.YELLOW.index);
        style7.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style7.setAlignment(HorizontalAlignment.FILL);
        cell.setCellValue("FOREGROUND COLOR/FOREGROUND PATTERN");
        cell.setCellStyle(style7);
        sheet.setColumnWidth(0,5000);        //sets width of the specified column
        
        //Write it to excel file
        try{
            FileOutputStream out = new FileOutputStream( new File("cellStyles.xlsx") );
            workbook.write(out);
            out.close();
            System.out.println("cellstyles.xlsx has been created successfully.");
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }
    
}
