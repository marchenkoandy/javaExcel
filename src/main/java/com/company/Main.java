package com.company;

import com.company.excel.ExcelWorkbook;
import com.company.frame.MainDialog;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.IOException;

public class Main {

    public static void main(String[] args){
//Create a file chooser
//        MainDialog dialog = new MainDialog();
//        dialog.pack();
//        dialog.setVisible(true);
////        System.exit(0);
//        final JFileChooser fc = new JFileChooser();
//        FileNameExtensionFilter filter = new FileNameExtensionFilter("JPG & GIF Images", "jpg", "gif");
//        fc.setFileFilter(filter);
//        //In response to a button click:
//        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
//
//        fc.showOpenDialog(dialog);
//        int returnVal = fc.showOpenDialog();
//        ExcelWorkbook excelWorkbook = new ExcelWorkbook();
//        excelWorkbook.setInputFile("C:/Temp/ACO_LMO/ACO.xlsx").read().print();
        MainDialog dialog = new MainDialog();
        dialog.pack();
        dialog.setVisible(true);
        System.exit(0);
    }
}
