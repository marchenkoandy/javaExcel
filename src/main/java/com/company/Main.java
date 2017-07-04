package com.company;

import com.company.excel.ExcelWorkbook;
import com.company.frame.MainDialog;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
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
        ExcelWorkbook excelWorkbook = new ExcelWorkbook();
        excelWorkbook.setInputFile("C:/Users/amarchenko/Desktop/Java_Excel/ACO.xls").read().print();
    }
}
