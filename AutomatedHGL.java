/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */


package automatedhgl;

/**
 *
 * @author ahudson and others
 */

import java.sql.ResultSet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class AutomatedHGL 
{
    
    static double InletStr;
    static double OutWSE;
    int Do = 1;
    int Qo = 1;
    int Lo = 1;
    String SFO;
        
    char Hf = 'G';
    int Vo = 1;
    char Ho = 'I';
    int Qi = 1;
    int Vi = 1;
    int QiVi = Qi * Vi;
    char Vi2g = 'M';
    char Hi = 'N';
    int ANG = 1;
    char Ha = 'P';
    char Ht = 'Q';
    char HHt = 'R';
    char DHt = 'S';
    char FH = 'T';
    char IWSE = 'U';
    int OpenELE = 1;
    int InvIn = 1;
    char SurfInFlo = 'Y';
    char K = 'Z';
    char RimMnWSE = 'A';
    char DIA = 'B';
    char InShape = 'C';
   
    
    static int count;
    
    
    
    private static final String INFILE_NAME = "/tmp/STORMDES.xlsx";
    
    private static final String FILE_NAME = "/tmp/HGLEXCEL.xlsx";
    
    public static void main(String[] args) 
    {
                            
        try 
        {
                        
            FileInputStream excelFile = new FileInputStream(new File(INFILE_NAME));
            
            //create workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
            
            //get first desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            
            //create workbook instance to output excel file
            XSSFWorkbook workbookHGL = new XSSFWorkbook();
            
            //create sheet in output excel file
            XSSFSheet sheetHGL = workbookHGL.createSheet("HGL");
            
            
            //iterate through each row one by one
            Iterator<Row> rowiterator = sheet.iterator();

            while (rowiterator.hasNext()) 
            {
                Row row = rowiterator.next();
                
                //for each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) 
                {

                    Cell cell = cellIterator.next();
                                                                                
                    if (row.getRowNum() > 7 && count < 23 )  //to filter column headings
                    {
                                                
                        //check the cell type and format accordingly
                        switch (cell.getCellType())
                        {
                            case Cell.CELL_TYPE_NUMERIC:
                                count++;
                                
                                //assign get value to correct variable
                                if(count == 1 ){InletStr = cell.getNumericCellValue();}
                                else if(count == 2 ){OutWSE = cell.getNumericCellValue();}
                                
                                System.out.print(cell.getNumericCellValue() + " (" + count + ") ");
                                break;
                                
                            case Cell.CELL_TYPE_STRING:
                                count++;
                                
                                /*//assign get value to correct variable
                                if( count == 1 ){InletStr = cell.getStringCellValue();}*/
                                
                                System.out.print(cell.getStringCellValue() + " (" + count + ") ");
                                break;
                                
                            case Cell.CELL_TYPE_FORMULA:
                                count++;
                                
                                /*//assign get value to correct variable
                                if( count == 1 ){InletStr = cell.getCachedFormulaResultType();}*/
                                
                                System.out.print(cell.getCachedFormulaResultType() + " (" + count + ") ");
                                break;
                        } 
                    }
                    
                    else
                    {
                        count = 0; //reset the count at the end of the row
                    }
                                       
                }
                                                
                System.out.println("return");
            }
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
                
        
        //Output Excel file
            
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
        Object[][] datatypes = {
                {"Datatype", "Type", "Size(in bytes)"},
                {"int", "Primitive", 2},
                {"float", "Primitive", 4},
                {"double", "Primitive", 8},
                {"char", "Primitive", 1},
                {"String", "Non-Primitive", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        System.out.print(InletStr + " ");
        System.out.print(OutWSE + " ");
        System.out.println("HGL Done");        
                
    }
    
}
