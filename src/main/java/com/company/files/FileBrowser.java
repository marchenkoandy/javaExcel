package com.company.files;

import com.company.excel.ExcelWorkbook;

import java.io.File;
import java.util.ArrayList;

/**
 * Created by AMarchenko on 7/4/2017.
 */
public class FileBrowser {
    private ArrayList<String> listOfFiles = new ArrayList<String>();
    private void verifyAndAddFile(File fileIn){
        String fileName = fileIn.getName();
        String fileExtensionName = fileName.substring(fileName.lastIndexOf("."));
        boolean b1 = fileExtensionName.equals(ExcelWorkbook.XLSX);
        boolean b2 = fileExtensionName.equals(ExcelWorkbook.XLS);
        boolean b3 = !fileName.contains(ExcelWorkbook.TILDA);
        if ((b1 || b2) && b3){
            listOfFiles.add(fileIn.getAbsolutePath());
        }
    }
    public void getFilesFromSingleFolder(File inputFolder){
        if (inputFolder.isDirectory()) {
            File[] listOfFilesInputFolder = inputFolder.listFiles();
            for (int i=0;i<listOfFilesInputFolder.length;i++){
                if (listOfFilesInputFolder[i].isFile()){
                    verifyAndAddFile(new File(listOfFilesInputFolder[i].getAbsolutePath()));
                }
                else
                {
                    getFilesFromSingleFolder(listOfFilesInputFolder[i]);
                }
            }
        }
        else
        {
            verifyAndAddFile(inputFolder);
        }
    }
    public ArrayList<String> recursiveListOfFiles(){
        return listOfFiles;
    }
}
