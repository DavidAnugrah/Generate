/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generatetotxt;

import id.my.berkah.util.DynamicQuery;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Vector;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 *
 * @author azec
 */
public class exeTXT {

    DynamicQuery dynamicQuery;
   

    List<String[]> data = new ArrayList<String[]>();

    public WritableSheet sheet1;
    public WritableCellFormat format;
    WritableWorkbook writebook;

    public void Query1(String query) {
        Connection conn = MyBatisConnectionFactory.getSqlSessionFactory().openSession().getConnection();

        SimpleDateFormat df = new SimpleDateFormat("ddMMyyyyHHmmss");
        Date date = new Date();
        String newDate = df.format(date);
        try {
            File fileOut = new File("E://" + "fileExport" + newDate + ".txt");
            fileOut.createNewFile();
            FileWriter fw = new FileWriter(fileOut.getAbsoluteFile());
            BufferedWriter bf = new BufferedWriter(fw);

            this.dynamicQuery = new DynamicQuery(conn);

            this.dynamicQuery.showData(query);
            for (String columnName : this.dynamicQuery.getColumnNames()) {
                System.out.println(columnName);
                bf.write(columnName);
                bf.write("      ");
            }
            bf.write("\r\n");
            int b = 0;
            for (Vector<Object> dataRow : this.dynamicQuery.getData()) {
                String[] dataRowArray = new String[dataRow.size()];
                int column = 0;

                int a = 0;

                for (Object object : dataRow) {
                    try {
                        dataRowArray[column] = object.toString();
                        System.out.println(object.toString());
                        bf.write(object.toString());
                        bf.write("            ");
                        Label row = new Label(a, b + 1, object.toString());
                    } catch (Exception e) {
                        dataRowArray[column] = "";
                    }
                    a++;
                    column++;
                }
                
                data.add(dataRowArray);
           bf.write("\r\n");
           bf.write(" ");
             b++;
            }
           bf.close();

        } catch (Exception e) {

        }

    }

}
