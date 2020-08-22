/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package seatarrangement;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import net.proteanit.sql.DbUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author vishnu
 */
public class ViewArrangement extends javax.swing.JPanel {
    public int ar[][];
    String room[];
    public int b;
    int[] bench=new int[12];
    /** Creates new form ViewArrangement */
    public ViewArrangement() {
        initComponents();
    }
    public  ViewArrangement(int a[][],String rooms[],int bench){
        ar=a;
        room=rooms;
        b=bench;
        table(b);
       
    }
    public void deleteArrangement(){
        String url ="jdbc:mysql://localhost/sas?autoReconnect=true&serverTimezone=UTC&useSSL=False&allowPublicKeyRetrieval=true";
            try (Connection con = DriverManager.getConnection(url, "root", "")){
                String sql="DELETE FROM arrangement";
                Statement st2=con.createStatement();
                st2.execute(sql);
            } catch (SQLException ex) {
            Logger.getLogger(ViewArrangement.class.getName()).log(Level.SEVERE, null, ex);
        }
            
    }
    public void table(int b1){
        deleteArrangement();
        String url ="jdbc:mysql://localhost/sas?autoReconnect=true&serverTimezone=UTC&useSSL=False&allowPublicKeyRetrieval=true";
            try (Connection con = DriverManager.getConnection(url, "root", "")){
                int[] rollno=new int[3];
                String[] name=new String[3];
                for(int i=0;i<12;i++){
                    String r=room[i];
                    String sql2="SELECT benches FROM room where roomid="+r+"";
                    Statement st2=con.createStatement();
                    ResultSet rs2=st2.executeQuery(sql2);
                    if(rs2.next()){
                        bench[i]=rs2.getInt(1);
                    }
                    else{
                        bench[i]=0;
                    }
                }
                for(int i=0;i<12;i++){
                    System.out.println(bench[i]);
                }
                System.out.println(b1);
                int k=0,l=0;
                for(int i=0;i<b;i++){
                    if(k==bench[l]){
                        k=0;
                        l++;
                    }
                    for(int j=0;j<3;j++){
                        String sql1="SELECT rollno,name FROM student where reg_no='"+ar[i][j]+"'";
                        Statement st1=con.createStatement();
                        ResultSet rs1=st1.executeQuery(sql1);
                        if(rs1.next()){
                            rollno[j]=rs1.getInt(1);
                            name[j]=rs1.getString(2);
                            System.out.println("name is"+name[j]);
                        }
                        else{
                            name[j]="Null";
                            rollno[j]=0;
                        }
                    }
                    String sql3="INSERT INTO `arrangement`(`room_no`, `c1_roll_no`, `c1_name`, `c2_roll_no`, `c2_name`, `c3_roll_no`, `c3_name`) VALUES (?,?,?,?,?,?,?)";
                    PreparedStatement pst=con.prepareStatement(sql3);
                    pst.setString(1,room[l]);
                    pst.setInt(2,rollno[0]);
                    pst.setString(3,name[0]);
                    pst.setInt(4,rollno[1]);
                    pst.setString(5,name[1]);
                    pst.setInt(6,rollno[2]);
                    pst.setString(7,name[2]);
                    pst.executeUpdate();
                    k++;
                }
            }catch (SQLException ex) {
            Logger.getLogger(ViewArrangement.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void printTable(){
        String url ="jdbc:mysql://localhost/sas?autoReconnect=true&serverTimezone=UTC&useSSL=False&allowPublicKeyRetrieval=true";
            try (Connection con = DriverManager.getConnection(url, "root", "")) {
                String sql="SELECT `room_no`, `c1_roll_no`, `c1_name`, `c2_roll_no`, `c2_name`, `c3_roll_no`, `c3_name` FROM `arrangement`";
                PreparedStatement pst=con.prepareStatement(sql);
                ResultSet resultSet=pst.executeQuery(sql);
                resultTable.setModel(DbUtils.resultSetToTableModel(resultSet));
            } catch (SQLException ex) {
            Logger.getLogger(Panel2.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        resultTable = new javax.swing.JTable();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        fileName = new javax.swing.JTextField();

        setBackground(new java.awt.Color(255, 255, 255));

        resultTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        jScrollPane1.setViewportView(resultTable);

        jButton1.setBackground(new java.awt.Color(255, 153, 0));
        jButton1.setFont(new java.awt.Font("SansSerif", 0, 18)); // NOI18N
        jButton1.setText("view");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setBackground(new java.awt.Color(255, 153, 0));
        jButton2.setFont(new java.awt.Font("SansSerif", 0, 18)); // NOI18N
        jButton2.setText("EXPORT");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("SansSerif", 0, 18)); // NOI18N
        jLabel1.setText("File Name");

        fileName.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        fileName.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileNameActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(35, 35, 35)
                        .addComponent(jScrollPane1))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(190, 190, 190)
                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 101, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(fileName, javax.swing.GroupLayout.PREFERRED_SIZE, 254, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(32, 32, 32)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 141, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(126, 126, 126)))
                .addGap(52, 52, 52))
            .addGroup(layout.createSequentialGroup()
                .addGap(385, 385, 385)
                .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(407, 407, 407))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGap(52, 52, 52)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 373, Short.MAX_VALUE)
                .addGap(27, 27, 27)
                .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 44, Short.MAX_VALUE)
                .addGap(38, 38, 38)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(4, 4, 4)
                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(4, 4, 4)
                        .addComponent(fileName))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(1, 1, 1)
                        .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, 43, Short.MAX_VALUE)))
                .addGap(29, 29, 29))
        );
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        printTable();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        String fname=fileName.getText();
    try {
        // TODO add your handling code here:
        XSSFWorkbook workbook=new XSSFWorkbook();
        FileOutputStream out= new FileOutputStream(new File("C:\\Users\\vishnu\\Pictures\\"+fname+".xlsx"));
        String url ="jdbc:mysql://localhost/sas?autoReconnect=true&serverTimezone=UTC&useSSL=False&allowPublicKeyRetrieval=true";
        Connection con = DriverManager.getConnection(url, "root", "");
        String sql="SELECT `room_no`, `c1_roll_no`, `c1_name`, `c2_roll_no`, `c2_name`, `c3_roll_no`, `c3_name` FROM `arrangement`";
        Statement st2=con.createStatement();
        ResultSet r=st2.executeQuery(sql);
        r.next();
        int flag=0,k=0;
        while(flag==0){
            String ro=r.getString(1);
            XSSFSheet sheet=workbook.createSheet("room"+ro);
            XSSFRow row=sheet.createRow(0);
            XSSFCell cell=row.createCell(0);
            cell.setCellValue("Rollno");
            cell=row.createCell(1);
            cell.setCellValue("Name");
            cell=row.createCell(2);
            cell.setCellValue("Sign");
            cell=row.createCell(3);
            cell.setCellValue("Rollno");
            cell=row.createCell(4);
            cell.setCellValue("Name");
            cell=row.createCell(5);
            cell.setCellValue("Sign");
            cell=row.createCell(6);
            cell.setCellValue("Rollno");
            cell=row.createCell(7);
            cell.setCellValue("Name");
            cell=row.createCell(8);
            cell.setCellValue("Sign");
            int i=1;
            while(ro.equalsIgnoreCase(r.getString(1))){
                row=sheet.createRow(i);
                if(r.getInt(2)!=0){
                    cell=row.createCell(0);
                    cell.setCellValue(r.getInt(2));
                    cell=row.createCell(1);
                    cell.setCellValue(r.getString(3));
                }
                if(r.getInt(4)!=0){
                    cell=row.createCell(3);
                    cell.setCellValue(r.getInt(4));
                    cell=row.createCell(4);
                    cell.setCellValue(r.getString(5));
                }
                if(r.getInt(6)!=0){
                    cell=row.createCell(6);
                    cell.setCellValue(r.getInt(6));
                    cell=row.createCell(7);
                    cell.setCellValue(r.getString(7));
                }
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(4);
                sheet.autoSizeColumn(7);
                if(r.next()){
                    flag=0;
                }
                else{
                    flag=1;
                    break;
                }
                i++;
            }
           k++;
        }
        workbook.write(out);
        out.close();
    } catch (FileNotFoundException ex) {
        Logger.getLogger(ViewArrangement.class.getName()).log(Level.SEVERE, null, ex);
    } catch (IOException ex) { 
        Logger.getLogger(ViewArrangement.class.getName()).log(Level.SEVERE, null, ex);
    } catch (SQLException ex) {
        Logger.getLogger(ViewArrangement.class.getName()).log(Level.SEVERE, null, ex);
    } 
    }//GEN-LAST:event_jButton2ActionPerformed

    private void fileNameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_fileNameActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_fileNameActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextField fileName;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable resultTable;
    // End of variables declaration//GEN-END:variables

}
