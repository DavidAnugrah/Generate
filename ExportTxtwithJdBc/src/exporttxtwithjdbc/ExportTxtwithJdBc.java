/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exporttxtwithjdbc;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 *
 * @author azec
 */
public class ExportTxtwithJdBc {
SimpleDateFormat df = new SimpleDateFormat("ddMMyyyyHHmmss");
    Date date = new Date();
    String newDate = df.format(date);

    private String url = "jdbc:oracle:thin:@192.168.3.33:2016:devl";
    private String driverClassName = "oracle.jdbc.driver.OracleDriver";
    private String username = "APPS";
    private String password = "11GAPPS";
//    private String path = "E://" + "export" + newDate + ".xls";

    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    
      boolean generate() {
         try {
                File file = new File("E://" + "export" + newDate + ".txt");
                    FileWriter fw = new FileWriter(file.getAbsoluteFile());
                    BufferedWriter bw = new BufferedWriter(fw);

            conn = getConnection();
            String query = "select user_id, user_name, user_description from fnd_users";

            pst = conn.prepareStatement(query);
            rs = pst.executeQuery();
            bw.write("\r\n");
            bw.close();   
                      }
         
                catch (Exception e) 
                {  
                    System.out.println(e);  
                } 
                finally 
                {  
                    if (conn != null)
                    {  
                        try {  
                            conn.close();  
                        } catch (SQLException e) {  
                            e.printStackTrace();  
                        }  
                    }  
                    if (pst != null) {  
                        try {  
                            pst.close();  
                        } catch (SQLException e) {  
                            e.printStackTrace();  
                        }  
                    }  
                    if (rs != null) {  
                        try {  
                            rs.close();  
                        } catch (SQLException e)
                        {  
                            e.printStackTrace();  
                        }  
                    }  
                }
    return true;
      }
   
    public Connection getConnection() {
        Connection con = null;
        try {
            Class.forName(driverClassName);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        try {
            con = DriverManager.getConnection(url, username, password);
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return con;
    }
    
    public static void main(String[] args) {
        // TODO code application logic here
        try {
            
              ExportTxtwithJdBc exportTxtwithJdBc = new ExportTxtwithJdBc();
        if (exportTxtwithJdBc.generate()) {
            
         System.out.println("Data exported..");
            } else {
                System.out.println("Data not exported..");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
}
}

    

