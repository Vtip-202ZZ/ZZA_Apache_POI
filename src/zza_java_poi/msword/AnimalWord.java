package zza_java_poi.msword;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;

public class AnimalWord extends javax.swing.JFrame {

    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread {

        @Override
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath() + System.getProperty("file.separator");

            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "animal_temp.doc")) {
                doc = new HWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Ошибка нахождения образа!");
            }

            try {
                doc.getRange().replaceText("$Зебра", jTextField_Zebra.getText());
                doc.getRange().replaceText("$Лев", jTextField_Lion.getText());
                doc.getRange().replaceText("$Жираф", jTextField_Giraffe.getText());
                doc.getRange().replaceText("$Бегемот", jTextField_Hippopotamus.getText());
                doc.getRange().replaceText("$Орёл", jTextField_Aquila.getText());
            } catch (Exception ex) {
                System.err.println("Ошибка замены текста!");
            }

            try (FileOutputStream fos = new FileOutputStream(dir + "animal.doc")) {
                doc.write(fos);
                fos.close();
                Desktop.getDesktop().open(new File(dir + "animal.doc"));
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    public AnimalWord() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField_Zebra = new javax.swing.JTextField();
        jTextField_Lion = new javax.swing.JTextField();
        jButton_export = new javax.swing.JButton();
        jTextField_Giraffe = new javax.swing.JTextField();
        jTextField_Hippopotamus = new javax.swing.JTextField();
        jTextField_Aquila = new javax.swing.JTextField();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Животные MS Word");
        setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        setResizable(false);
        getContentPane().setLayout(null);
        getContentPane().add(jTextField_Zebra);
        jTextField_Zebra.setBounds(150, 140, 120, 30);
        getContentPane().add(jTextField_Lion);
        jTextField_Lion.setBounds(150, 95, 120, 30);

        jButton_export.setText("Экспортировать в Word");
        jButton_export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_exportActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_export);
        jButton_export.setBounds(290, 303, 180, 30);
        getContentPane().add(jTextField_Giraffe);
        jTextField_Giraffe.setBounds(150, 180, 120, 30);

        jTextField_Hippopotamus.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_HippopotamusActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Hippopotamus);
        jTextField_Hippopotamus.setBounds(150, 230, 120, 30);
        getContentPane().add(jTextField_Aquila);
        jTextField_Aquila.setBounds(150, 270, 120, 25);

        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/zza_java_poi/msword/animal.PNG"))); // NOI18N
        getContentPane().add(jLabel1);
        jLabel1.setBounds(0, 0, 730, 340);

        setBounds(0, 0, 744, 376);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_exportActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton_exportActionPerformed

    private void jTextField_HippopotamusActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_HippopotamusActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_HippopotamusActionPerformed

    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(AnimalWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }

        java.awt.EventQueue.invokeLater(() -> {
            new AnimalWord().setVisible(true);
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_export;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JTextField jTextField_Aquila;
    private javax.swing.JTextField jTextField_Giraffe;
    private javax.swing.JTextField jTextField_Hippopotamus;
    private javax.swing.JTextField jTextField_Lion;
    private javax.swing.JTextField jTextField_Zebra;
    // End of variables declaration//GEN-END:variables
}
