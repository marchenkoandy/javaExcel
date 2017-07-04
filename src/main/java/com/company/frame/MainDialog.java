package com.company.frame;

import com.company.excel.ExcelWorkbook;
import com.company.excel.Result;

import javax.swing.*;
import java.awt.event.*;
import java.util.ArrayList;

public class MainDialog extends JDialog {
    private JPanel contentPane;
    private JButton buttonFile;
    private JButton buttonCancel;
    private JButton buttonFolder;

    private JScrollPane textArea;
    private JTextArea textArea1;
    private ArrayList<Result> results;
    public MainDialog() {
        setContentPane(contentPane);
        setModal(true);
        getRootPane().setDefaultButton(buttonFile);

        buttonFile.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onFile();
            }
        });

        buttonCancel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        });

        // call onCancel() when cross is clicked
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                onCancel();
            }
        });

        // call onCancel() on ESCAPE
        contentPane.registerKeyboardAction(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onCancel();
            }
        }, KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);
        buttonFolder.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onFolder();
            }
        });

    }

    private void onFolder() {
        // add your code here
        dispose();
    }
    private void onFile() {
        // add your code here
        ExcelWorkbook excelWorkbook = new ExcelWorkbook();
        results = excelWorkbook.setInputFile("C:/Temp/ACO_LMO/ACO.xlsx").read();
        for (Result r:results) {
            String line;
            line = "%s  %s  %s  %s  %s";
            line = String.format(line, r.cellValue, r.cellTypeName, r.cellSheetName, r.cellWorkbookName, r.cellColumnNumber);
            textArea1.append(line + "\n");
        }
//            System.out.println(line);
//        textArea1.append(results.get(0).cellValue);
//        dispose();
    }

    private void onCancel() {
        // add your code here if necessary
        dispose();
    }

    public static void main(String[] args) {
        MainDialog dialog = new MainDialog();
        dialog.pack();
        dialog.setVisible(true);
        System.exit(0);
    }
}
