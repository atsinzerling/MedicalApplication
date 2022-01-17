/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package MedicalApplication;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static MedicalApplication.MainFrame.NumRows;
import static MedicalApplication.MainFrame.headers_adding;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFRow;
import static MedicalApplication.MainFrame.NumGrades;

//This class adds a new student to the excel file

public class AddStudentFrame extends javax.swing.JFrame {

    /**
     * Creates new form SecondFrame
     */
    
    public static ArrayList < String > ArrListNames = new ArrayList < String > ();
public static ArrayList < String > ArrListGrade = new ArrayList < String > ();
public static ArrayList < String > ArrListExemptedFrom = new ArrayList < String > ();
public static ArrayList < String > ArrListExemptedTo = new ArrayList < String > ();
public static ArrayList < String > ArrListComments = new ArrayList < String > ();

public static String[] GradesArray = new String[NumGrades];

public static String[] NewStudent = {
    "-",
    "-",
    "-",
    "-",
    "-"
};

//This method loads the excel file and creates an array for GradeCheckBox

public AddStudentFrame() {
    if (NumGrades != GradesArray.length){
        GradesArray = new String[NumGrades];
    }


    try {
        File ExcelFile = new File(System.getProperty("user.dir") + "/src/Data/teachers.xlsx");

        if (ExcelFile.exists() == false) {
            for (int i = 0; i < GradesArray.length; i++) {
                GradesArray[i] = "";
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
                //Iterate through each rows from first sheet
                Iterator < Row > rowIterator = sheet.iterator();
                Row row = rowIterator.next();
                int rows = 0;
                while (rowIterator.hasNext()) {
                    row = rowIterator.next();

                    Iterator < Cell > cellIterator = row.cellIterator();

                    Cell cell = cellIterator.next();
                    GradesArray[rows] = cell.getStringCellValue();
                    rows++;
                }

            }
            Workbook.close();
            BIS.close();


        }
    } catch (FileNotFoundException ex) {
        Logger.getLogger(AddStudentFrame.class.getName()).log(Level.SEVERE, null, ex);
    } catch (IOException ex) {
        Logger.getLogger(AddStudentFrame.class.getName()).log(Level.SEVERE, null, ex);
    } 
//    catch () {
//    }

    initComponents();
    NumRows++;

    try {
        File ExcelFile = new File(System.getProperty("user.dir") + "/src/Data/table.xlsx");
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
            NumRows = 1;
            Row row = rowIterator.next();
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                Iterator < Cell > cellIterator = row.cellIterator();
                String strr = new String();
                strr = cellIterator.next().getStringCellValue();
                if (strr.equals("")) {
                    ArrListNames.add(" ");
                } else {
                    ArrListNames.add(strr);
                }
                strr = cellIterator.next().getStringCellValue();

                if (strr.equals("")) {
                    ArrListGrade.add(" ");
                } else {
                    ArrListGrade.add(strr);
                }
                strr = cellIterator.next().getStringCellValue();

                if (strr.equals("")) {
                    ArrListExemptedFrom.add(" ");
                } else {
                    ArrListExemptedFrom.add(strr);
                }
                strr = cellIterator.next().getStringCellValue();

                if (strr.equals("")) {
                    ArrListExemptedTo.add(" ");
                } else {
                    ArrListExemptedTo.add(strr);
                }
                strr = cellIterator.next().getStringCellValue();

                if (strr.equals("")) {
                    ArrListComments.add(" ");
                } else {
                    ArrListComments.add(strr);
                }
                NumRows++;

            }

        }
        Workbook.close();
        BIS.close();

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

        jPanel1 = new javax.swing.JPanel();
        Name = new javax.swing.JLabel();
        Grade = new javax.swing.JLabel();
        ExemptedFrom = new javax.swing.JLabel();
        ExemptedTo = new javax.swing.JLabel();
        Comments = new javax.swing.JLabel();
        NameField = new javax.swing.JTextField();
        GradeComboBox = new javax.swing.JComboBox<>();
        ExemptedFromField = new javax.swing.JTextField();
        ExemptedToField = new javax.swing.JTextField();
        CommentScrollPane = new javax.swing.JScrollPane();
        CommentArea = new javax.swing.JTextArea();
        RirectlySendLetterButton = new javax.swing.JButton();
        EditLetterButton = new javax.swing.JButton();
        SaveButton = new javax.swing.JButton();
        BackButton = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setMinimumSize(new java.awt.Dimension(800, 500));
        getContentPane().setLayout(null);

        jPanel1.setLayout(null);

        Name.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        Name.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        Name.setText("Full name of the student:");
        jPanel1.add(Name);
        Name.setBounds(140, 40, 150, 20);

        Grade.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        Grade.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        Grade.setText("Grade:");
        jPanel1.add(Grade);
        Grade.setBounds(160, 70, 130, 20);

        ExemptedFrom.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        ExemptedFrom.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        ExemptedFrom.setText("Exempted from date:");
        jPanel1.add(ExemptedFrom);
        ExemptedFrom.setBounds(150, 100, 140, 20);

        ExemptedTo.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        ExemptedTo.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        ExemptedTo.setText("Exempted to date:");
        jPanel1.add(ExemptedTo);
        ExemptedTo.setBounds(150, 130, 140, 20);

        Comments.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        Comments.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        Comments.setText("Comments:");
        jPanel1.add(Comments);
        Comments.setBounds(180, 160, 110, 20);
        jPanel1.add(NameField);
        NameField.setBounds(310, 40, 240, 20);

        GradeComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(GradesArray));
        GradeComboBox.setMinimumSize(new java.awt.Dimension(200, 19));
        GradeComboBox.setPreferredSize(new java.awt.Dimension(400, 19));
        jPanel1.add(GradeComboBox);
        GradeComboBox.setBounds(310, 70, 80, 19);
        GradeComboBox.getAccessibleContext().setAccessibleName("");

        jPanel1.add(ExemptedFromField);
        ExemptedFromField.setBounds(310, 100, 130, 20);
        jPanel1.add(ExemptedToField);
        ExemptedToField.setBounds(310, 130, 130, 20);

        CommentArea.setColumns(20);
        CommentArea.setRows(5);
        CommentScrollPane.setViewportView(CommentArea);

        jPanel1.add(CommentScrollPane);
        CommentScrollPane.setBounds(310, 160, 240, 100);

        RirectlySendLetterButton.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        RirectlySendLetterButton.setText("Directly send the letter to the teacher");
        RirectlySendLetterButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                RirectlySendLetterButtonActionPerformed(evt);
            }
        });
        jPanel1.add(RirectlySendLetterButton);
        RirectlySendLetterButton.setBounds(500, 310, 230, 40);

        EditLetterButton.setText("Edit the letter");
        EditLetterButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                EditLetterButtonActionPerformed(evt);
            }
        });
        jPanel1.add(EditLetterButton);
        EditLetterButton.setBounds(500, 360, 170, 30);

        SaveButton.setText("Save");
        SaveButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveButtonActionPerformed(evt);
            }
        });
        jPanel1.add(SaveButton);
        SaveButton.setBounds(310, 270, 120, 23);

        BackButton.setText("Back");
        BackButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                BackButtonActionPerformed(evt);
            }
        });
        jPanel1.add(BackButton);
        BackButton.setBounds(10, 30, 90, 40);

        getContentPane().add(jPanel1);
        jPanel1.setBounds(0, 0, 800, 500);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    //This method saves the data fro EditMailFrame and opens it 
    
    private void EditLetterButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_EditLetterButtonActionPerformed
        try {
    try (BufferedWriter writer = new BufferedWriter(new FileWriter(System.getProperty("user.dir") + "/src/Data/SendData.txt"))) {
        writer.write(NameField.getText() + "\n" + GradeComboBox.getSelectedIndex() + "\n" + ExemptedFromField.getText() + "\n" + ExemptedToField.getText() + "\n" + CommentArea.getText());
    }
} catch (IOException ex) {
    Logger.getLogger(SettingsFrame.class.getName()).log(Level.SEVERE, null, ex);
}

try {
    EditMailFrame EditMail = new EditMailFrame();
    EditMail.setVisible(true);
} catch (IOException ex) {
    Logger.getLogger(AddStudentFrame.class.getName()).log(Level.SEVERE, null, ex);
}
    }//GEN-LAST:event_EditLetterButtonActionPerformed

    //This method  saves data of the new student
    
    private void SaveButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveButtonActionPerformed
        NewStudent[0] = NameField.getText();
        NewStudent[1] = GradeComboBox.getSelectedItem().toString();
        NewStudent[2] = ExemptedFromField.getText();
        NewStudent[3] = ExemptedToField.getText();
        NewStudent[4] = CommentArea.getText();
                
    }//GEN-LAST:event_SaveButtonActionPerformed

    //This method returns the user bak to the MainFrame
    
    private void BackButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_BackButtonActionPerformed
        if ((NewStudent[0] != "-") && (NewStudent[1] != "-") && (NewStudent[2] != "-") && (NewStudent[3] != "-") && (NewStudent[4] != "-")) {
        try {

            ArrListNames.add(0, NewStudent[0]);
            ArrListGrade.add(0, NewStudent[1]);
            ArrListExemptedFrom.add(0, NewStudent[2]);
            ArrListExemptedTo.add(0, NewStudent[3]);
            ArrListComments.add(0, NewStudent[4]);
            String[][] data = new String[NumRows][5];
            for (int i = 0; i < NumRows; i++) {
                data[i][0] = ArrListNames.get(i);
                data[i][1] = ArrListGrade.get(i);
                data[i][2] = ArrListExemptedFrom.get(i);
                data[i][3] = ArrListExemptedTo.get(i);
                data[i][4] = ArrListComments.get(i);
            }

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            Row row = sheet.createRow((short) 0);
            for (int i = 0; i < 5; i++) {
                Cell cell = row.createCell((short) i);
                cell.setCellValue(headers_adding[i]);
            }


            for (int i = 0; i < ArrListNames.size(); i++) {
                row = sheet.createRow((short) i + 1);

                Cell cell = row.createCell((short) 0);
                cell.setCellValue(ArrListNames.get(i));

                cell = row.createCell((short) 1);
                cell.setCellValue(ArrListGrade.get(i));

                cell = row.createCell((short) 2);
                cell.setCellValue(ArrListExemptedFrom.get(i));

                cell = row.createCell((short) 3);
                cell.setCellValue(ArrListExemptedTo.get(i));

                cell = row.createCell((short) 4);
                cell.setCellValue(ArrListComments.get(i));
            }
            try (FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir") + "/src/Data/table.xlsx")) {
                wb.write(fileOut);
            }
        } catch (IOException e) {}

    }



    ArrListNames.clear();
    ArrListGrade.clear();
    ArrListExemptedFrom.clear();
    ArrListExemptedTo.clear();
    ArrListComments.clear();
    for (int i = 0; i < 5; i++) {
        NewStudent[i] = "-";
    }
    dispose();
    MainFrame JFrameObj = new MainFrame();
    JFrameObj.setVisible(true);
    }//GEN-LAST:event_BackButtonActionPerformed

    //This method sends the letter without opening a new frame
    
    private void RirectlySendLetterButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_RirectlySendLetterButtonActionPerformed
        try {

        String StudentName = NameField.getText();
        int StudentGradeIndex = GradeComboBox.getSelectedIndex();
        String ExemptedFrom = ExemptedFromField.getText();
        String ExemptedTo = ExemptedToField.getText();
        String Comments = CommentArea.getText();


        File ExcelFile = new File(System.getProperty("user.dir") + "/src/Data/teachers.xlsx");
        FileInputStream FIS = new FileInputStream(ExcelFile);
        BufferedInputStream BIS = new BufferedInputStream(FIS);

        //Get the workbook instance for XLS file
        XSSFWorkbook Workbook = new XSSFWorkbook(BIS);

        //Get first sheet from the workbook
        XSSFSheet sheet = Workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(StudentGradeIndex + 1);

        Iterator < Cell > cellIterator = row.cellIterator();
        //Iterate through each rows from first sheet

        String StudentGrade = cellIterator.next().getStringCellValue();
        String TeacherName = cellIterator.next().getStringCellValue();
        String TeacherEmail1 = cellIterator.next().getStringCellValue();
        String TeacherEmail2 = cellIterator.next().getStringCellValue();


        String Email = (TeacherEmail1 + "\n" + TeacherEmail2);
        String Subject = "Exemption of " + StudentName;
        String Message = "Dear " + TeacherName + "," + "\n" + StudentName + " from " + StudentGrade + " is exempted from lessons of Physical Eduction from " + ExemptedFrom + " to " + ExemptedTo + "\n" + Comments + "\n" + "Medical department of Skolkovo Gymnasium";

        Workbook.close();
        FIS.close();
        BIS.close();
        BufferedWriter writer = new BufferedWriter(new FileWriter(System.getProperty("user.dir") + "/src/Data/SendData.txt"));
        writer.close();

        String[] lines = Email.split("\\n");

        BufferedReader reader = new BufferedReader(new FileReader(System.getProperty("user.dir") + "/src/Data/Email.txt"));
        String User = reader.readLine();
        String Password = reader.readLine();

        reader.close();

        for (int j = 0; j < lines.length; j++) {
            SendMail.send(lines[j], Subject, Message, User, Password);
        }

    } catch (IOException e) {}
    }//GEN-LAST:event_RirectlySendLetterButtonActionPerformed

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
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(AddStudentFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(AddStudentFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(AddStudentFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(AddStudentFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        //</editor-fold>

        /* Create and display the form */
        
        
        java.awt.EventQueue.invokeLater(() -> {
            new AddStudentFrame().setVisible(true);
        });
        
        
        
    }
    
    

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton BackButton;
    private javax.swing.JTextArea CommentArea;
    private javax.swing.JScrollPane CommentScrollPane;
    private javax.swing.JLabel Comments;
    private javax.swing.JButton EditLetterButton;
    private javax.swing.JLabel ExemptedFrom;
    private javax.swing.JTextField ExemptedFromField;
    private javax.swing.JLabel ExemptedTo;
    private javax.swing.JTextField ExemptedToField;
    private javax.swing.JLabel Grade;
    private javax.swing.JComboBox<String> GradeComboBox;
    private javax.swing.JLabel Name;
    private javax.swing.JTextField NameField;
    private javax.swing.JButton RirectlySendLetterButton;
    private javax.swing.JButton SaveButton;
    private javax.swing.JPanel jPanel1;
    // End of variables declaration//GEN-END:variables
}









