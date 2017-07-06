package com.company.frame;

//import com.company.excel.ExcelWorkbook_OLD;
import com.company.records.Result;

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
    private JLabel lableInfo;
    private JLabel lableProcessing;
    private ArrayList<Result> results;
    public MainDialog() {
        setContentPane(contentPane);
        setModal(true);
        getRootPane().setDefaultButton(buttonCancel);
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        buttonFolder.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                onFolder();
            }
        });
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


        pack();

        setSize(400,200);

    }

    private void onFolder() {
        // add your code here
        dispose();
    }
    private void onFile() {
        // add your code here

        dispose();
    }

    private void onCancel() {
        // add your code here if necessary
        dispose();
    }

    public static void main(String[] args) {
        MainDialog dialog = new MainDialog();
        dialog.pack();
        dialog.setSize(400,300);
        dialog.setVisible(true);
        System.exit(0);
    }

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }
}
