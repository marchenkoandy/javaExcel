package com.company.excel;

import com.company.Main;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorkbook {
    public static final String XLSX = ".xlsx";
    public static final String XLS = ".xls";
    public static final String TILDA = "~$";
    public static int LEVEL = Main.LEVEL + 1;
    public String name;
    public Workbook workbook = null;
    public ArrayList<ExcelWorksheet> sheets = new ArrayList<ExcelWorksheet>();
    public File fileExcel;

    public String getExtension() {
        return extension;
    }

    public void setExtension(String extension) {
        this.extension = extension;
    }

    private String extension;
    public ExcelWorkbook                (File fileExcel) {
        this.fileExcel      = fileExcel;
        this.name           = fileExcel.getAbsolutePath();
        String fileName     = this.name;
        this.setExtension(fileName.substring(fileName.lastIndexOf(".")));
    }

    public void                         info(){
        Main.printOnLevel(LEVEL,Main.delimiter());
        Main.printOnLevel(LEVEL,"Workbook file name:     " + name);
        Main.printOnLevel(LEVEL,"Workbook sheets count:  " + sheets.size());
        Main.printOnLevel(LEVEL,Main.delimiter());
    }
    public void                         printResults() {
        info();
        for (ExcelWorksheet sheet:sheets) {
            sheet.printResults();
        }
    }
    public void                         getData(){
        System.out.println("GETTING DATA " + name);
        for (int i = 0; i< workbook.getNumberOfSheets(); i++) {
            ExcelWorksheet worksheet = new ExcelWorksheet(workbook.getSheetAt(i));
            sheets.add(worksheet.getData());
        }
        this.close();
    }
    public ExcelWorkbook                read(){
        System.out.println("READING " + name);
        FileInputStream inputStream=null;
        try{
            inputStream= new FileInputStream(fileExcel);
            if (getExtension().equals(XLSX)) {
                workbook =new XSSFWorkbook(inputStream);
            }
            else if (getExtension().equals(XLS)) {
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
        return this;
    }

    public void                         write(){
        System.out.println("WRITING " + name);
        FileOutputStream outputStream = null;
        try{
            outputStream = new FileOutputStream(fileExcel);
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
    public void                         close() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (Exception e) {
                throw new RuntimeException();
            } finally {
                workbook = null;
                System.out.println("CLOSING " + name);
            }
        }
    }
}
