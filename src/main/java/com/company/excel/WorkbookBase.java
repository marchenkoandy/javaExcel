package com.company.excel;

import com.company.Main;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by AMarchenko on 7/7/2017.
 */
public class WorkbookBase {
    public static final String  XLSX = ".xlsx";
    public static final String  XLS = ".xls";
    public static final String  TILDA = "~$";


    public String               name;
    public String               extension;
    public File                 fileExcel;
    public Workbook             workbook;

    public                      WorkbookBase(File fileExcel){
        this.fileExcel              = fileExcel;
        this.name                   = fileExcel.getAbsolutePath();
        String fileName             = this.name;
        this.extension              = fileName.substring(fileName.lastIndexOf("."));
    }
    public void                 getWorkbook(){
        if (workbook == null) {
            if (this.fileExcel.exists()) {
                read();
            } else {
                create();
            }
        }
    }
    public void                 create() {
        if (extension.equals(XLSX)) {
            workbook = new XSSFWorkbook();
        } else if (extension.equals(XLS)) {
            workbook = new HSSFWorkbook();
        }

    }
    public void                 read(){
        Main.printOnLevel(Main.DEBUG_LEVEL,"READING FILE...: " + name);
        FileInputStream inputStream=null;
        try{
            inputStream= new FileInputStream(fileExcel);
            if (extension.equals(XLSX)) {
                workbook =new XSSFWorkbook(inputStream);
            }
            else if (extension.equals(XLS)) {
                workbook =new HSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
    public void                 write(){
        Main.printOnLevel(Main.DEBUG_LEVEL,"WRITING...: " + name);
        FileOutputStream outputStream = null;
        try{
            outputStream = new FileOutputStream(fileExcel);
            if (workbook.getNumberOfSheets()==0) {
                workbook.createSheet("Sheet1");
            }
            workbook.write(outputStream);
        }
        catch (Exception e){
            throw new RuntimeException();
        }
        finally {
            if (outputStream != null){
                try {
                    outputStream.flush();
                    outputStream.close();
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }
    }
    public void                 close() {
        Main.printOnLevel(Main.DEBUG_LEVEL,"CLOSING FILE...: " + name);
        if (workbook != null) {
            try {
                workbook.close();
            } catch (Exception e) {
                throw new RuntimeException();
            } finally {
                workbook = null;
            }
        }
    }
    public static String        convertDataToString(Cell cell){
        String sOut = "";
        if (cell != null) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    sOut = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    sOut = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    sOut = Double.toString(cell.getNumericCellValue());
                    break;
            }
        }
        return sOut;
    }
}
