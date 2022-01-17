/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package MedicalApplication;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static MedicalApplication.AddStudentFrame.GradesArray;

//This class contains the table of students that is shown to the user, functionality to import and export the table, and buttons leading to other frames.

public class MainFrame extends javax.swing.JFrame {
/**
 * Creates new form JFrame
 */
public static int NumRows;
public static int NumGrades;
public static String[] headers_adding = {
    "Full name",
    "Grade",
    "Exempted from date",
    "Exempted to date",
    "Comments"
};

//This method loads excel file, previously saved

public MainFrame() {
    initComponents();

    try {
        
        //File file = new File(System.getProperty("user.dir") + "/src/data");
        //file.mkdir();
//        System.out.println(System.getProperty("user.dir"));
        File dir = new File(System.getProperty("user.dir") + "/src/data");
        dir.mkdirs();
        
        File ExcelFile = new File(System.getProperty("user.dir") + "/src/data/table.xlsx");
        if (ExcelFile.exists() == false) {

            String[][] data;
            try (
                XSSFWorkbook Workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = Workbook.createSheet("new sheet");
                // Create a row and put some cells in it. Rows are 0 based.
                Row row = sheet.createRow((short) 0);
                data = null;
                for (int i = 0; i < headers_adding.length; i++) {
                    row.createCell(i).setCellValue(headers_adding[i]);
                } // Write the output to a file
                try (FileOutputStream FileOut = new FileOutputStream(System.getProperty("user.dir") + "/src/data/table.xlsx")) {
                    Workbook.write(FileOut);
                }
            }
            NumRows = 0;

            DefaultTableModel Model = (DefaultTableModel) StudentTable.getModel();
            Model.setDataVector(data, headers_adding);

        } else {

            BufferedInputStream BIS;
            try (FileInputStream FIS = new FileInputStream(ExcelFile)) {
                BIS = new BufferedInputStream(FIS);
                //Get first sheet from the workbook
                try ( //Get the workbook instance for XLS file
                    XSSFWorkbook Workbook = new XSSFWorkbook(BIS)) {
                    //Get first sheet from the workbook
                    XSSFSheet sheet = Workbook.getSheetAt(0);
                    //Iterate through each rows from first sheet
                    Iterator < Row > rowIterator = sheet.iterator();
                    Row row = rowIterator.next();
                    //JOptionPane.showMessageDialog(null, "Imported Successfully");
                    int cells = 0;
                    NumRows = sheet.getLastRowNum();
                    String[][] data = new String[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
                    int rows = 0;
                    while (rowIterator.hasNext()) {
                        row = rowIterator.next();

                        Iterator < Cell > cellIterator = row.cellIterator();

                        while (cellIterator.hasNext()) {

                            Cell cell = cellIterator.next();

                            data[rows][cells] = cell.getStringCellValue();

                            cells++;
                        }
                        rows++;
                        cells = 0;
                    }
                    DefaultTableModel Model = (DefaultTableModel) StudentTable.getModel();
                    Model.setDataVector(data, headers_adding);
                    Workbook.close();
                    FIS.close();
                }
            }
            BIS.close();

        }

    } catch (FileNotFoundException e) {} catch (IOException e) {}


    try {
        File ExcelFile = new File(System.getProperty("user.dir") + "/src/data/teachers.xlsx");

        if (ExcelFile.exists() == false) {

            for (int i = 0; i < GradesArray.length; i++) {
                NumGrades = 1;
            }
        } else {

            BufferedInputStream BIS;
            XSSFWorkbook Workbook;
            try (FileInputStream FIS = new FileInputStream(ExcelFile)) {
                BIS = new BufferedInputStream(FIS);
                //Get the workbook instance for XLS file
                Workbook = new XSSFWorkbook(BIS);
                //Get first sheet from the workbook
                XSSFSheet sheet = Workbook.getSheetAt(0);
                NumGrades = sheet.getLastRowNum();


            }
            Workbook.close();
            BIS.close();

        }
    } catch (FileNotFoundException ex) {
        Logger.getLogger(AddStudentFrame.class.getName()).log(Level.SEVERE, null, ex);
    } catch (IOException ex) {
        Logger.getLogger(AddStudentFrame.class.getName()).log(Level.SEVERE, null, ex);
    }


}

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        StudetTablePane = new javax.swing.JScrollPane();
        StudentTable = new javax.swing.JTable();
        ImportButton = new javax.swing.JButton();
        AddStudentButton = new javax.swing.JButton();
        SettingsButton = new javax.swing.JButton();
        ExportButton = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new java.awt.Dimension(800, 500));

        StudentTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        StudetTablePane.setViewportView(StudentTable);

        ImportButton.setText("Import new table");
        ImportButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ImportButtonActionPerformed(evt);
            }
        });

        AddStudentButton.setText("Add new student");
        AddStudentButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                AddStudentButtonActionPerformed(evt);
            }
        });

        SettingsButton.setText("Settings");
        SettingsButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SettingsButtonActionPerformed(evt);
            }
        });

        ExportButton.setText("Export the table");
        ExportButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ExportButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(30, 30, 30)
                .addComponent(AddStudentButton, javax.swing.GroupLayout.PREFERRED_SIZE, 150, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(200, 200, 200)
                .addComponent(ImportButton, javax.swing.GroupLayout.PREFERRED_SIZE, 140, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(10, 10, 10)
                .addComponent(ExportButton, javax.swing.GroupLayout.PREFERRED_SIZE, 120, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(20, 20, 20)
                .addComponent(SettingsButton, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGroup(layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addComponent(StudetTablePane, javax.swing.GroupLayout.PREFERRED_SIZE, 760, javax.swing.GroupLayout.PREFERRED_SIZE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(AddStudentButton, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(ImportButton, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(ExportButton, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(SettingsButton, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(10, 10, 10)
                .addComponent(StudetTablePane, javax.swing.GroupLayout.PREFERRED_SIZE, 390, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    //This method imports a new excel file and loads it
    
    private void ImportButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ImportButtonActionPerformed
        try {
    JFileChooser ExcelFileChooser = new JFileChooser("D:\\Mine stuff\\IA\\CS");
    ExcelFileChooser.setDialogTitle("Select Excel File");
    FileNameExtensionFilter Filter = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm");
    ExcelFileChooser.setFileFilter(Filter);
    int ExcelChooser = ExcelFileChooser.showOpenDialog(null);

    File ExcelFile = ExcelFileChooser.getSelectedFile();
    BufferedInputStream BIS;
    XSSFWorkbook Workbook;
    try (FileInputStream FIS = new FileInputStream(ExcelFile)) {
        BIS = new BufferedInputStream(FIS);
        //Get the workbook instance for XLS file
        Workbook = new XSSFWorkbook(BIS);
        //Get first sheet from the workbook
        XSSFSheet sheet = Workbook.getSheetAt(0);
        //Iterate through each rows from first sheet
        Iterator < Row > rowIterator = sheet.iterator();
        Row row = rowIterator.next();
        JOptionPane.showMessageDialog(null, "Imported Successfully");
        int cells = 0;
        NumRows = sheet.getLastRowNum();
        String[][] data = new String[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
        int rows = 0;
        while (rowIterator.hasNext()) {
            row = rowIterator.next();

            Iterator < Cell > cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                data[rows][cells] = cell.getStringCellValue();

                cells++;
            }
            rows++;
            cells = 0;
        }
        DefaultTableModel Model = (DefaultTableModel) StudentTable.getModel();
        Model.setDataVector(data, headers_adding);
        try (FileOutputStream FOS = new FileOutputStream(new File(System.getProperty("user.dir") + "/src/data/table.xlsx"))) {
            Workbook.write(FOS);
        }
    }
    BIS.close();
    Workbook.close();


} catch (FileNotFoundException e) {} catch (IOException e) {}
        
        
    }//GEN-LAST:event_ImportButtonActionPerformed

    //This method opens frame AAddStaudentFrame
    
    private void AddStudentButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_AddStudentButtonActionPerformed
        AddStudentFrame SecFrameObj = new AddStudentFrame();
        SecFrameObj.setVisible(true);
        dispose();
    }//GEN-LAST:event_AddStudentButtonActionPerformed

    //This method opens frame SettingsFrame
    
    private void SettingsButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SettingsButtonActionPerformed
        SettingsFrame SettingsFrame = new SettingsFrame();
        SettingsFrame.setVisible(true);
        dispose();
    }//GEN-LAST:event_SettingsButtonActionPerformed

    //This method exports the excel file
    
    private void ExportButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ExportButtonActionPerformed
        try {
    JFileChooser fileChooser = new JFileChooser();
    fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
    fileChooser.setSelectedFile(new File("fileToSave.txt"));
    int option = fileChooser.showSaveDialog(this);
    if (option == JFileChooser.APPROVE_OPTION) {

        String File = JOptionPane.showInputDialog(null, "Enter file name", "File name", JOptionPane.QUESTION_MESSAGE, null, null, "table.xlsx").toString();
        String FileName = fileChooser.getSelectedFile().getAbsolutePath() + "\\" + File;

        File ExcelFile = new File(System.getProperty("user.dir") + "/src/data/table.xlsx");
        FileInputStream FIS = new FileInputStream(ExcelFile);
        BufferedInputStream BIS = new BufferedInputStream(FIS);
        Workbook Workbook = new XSSFWorkbook(BIS);

        FileOutputStream FileOut = new FileOutputStream(FileName);
        Workbook.write(FileOut);
        FileOut.close();
        Workbook.close();
        JOptionPane.showMessageDialog(this, "The file was successfully exported", "File exported", JOptionPane.PLAIN_MESSAGE);

    }

} catch (Exception ex) {
    JOptionPane.showMessageDialog(this, "The file wasn't exported", "Try again", JOptionPane.WARNING_MESSAGE);
}
    }//GEN-LAST:event_ExportButtonActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new MainFrame().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton AddStudentButton;
    private javax.swing.JButton ExportButton;
    private javax.swing.JButton ImportButton;
    private javax.swing.JButton SettingsButton;
    private javax.swing.JTable StudentTable;
    private javax.swing.JScrollPane StudetTablePane;
    // End of variables declaration//GEN-END:variables
}
