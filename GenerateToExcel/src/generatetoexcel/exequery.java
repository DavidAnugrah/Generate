/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generatetoexcel;

import id.my.berkah.util.DynamicQuery;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 *
 * @author azec
 */
public class exequery {
    
     DynamicQuery dynamicQuery;
  private String title;
  private String queryTotal;
  private String queryTotalTemp;
  private String query;
  private String queryTemp;
  private String[] hflexColumn = new String[0];
  private int[] hiddenColumn = new int[0];
  private int[] selectedColumn = new int[0];
  private int rowSelectedIndex = -1;
  private String width = "1000";
  private String height = "500";
  private int startPageNumber = 0;
  private int pageSize = 10;
  private int param1;
  private int param2;
  private int total = 0;
 List<String[]> data = new ArrayList<String[]>();

    public WritableSheet sheet1;
    public WritableCellFormat format;
    WritableWorkbook writebook;
  
  public void Query1(String query){
        Connection conn = MyBatisConnectionFactory.getSqlSessionFactory().openSession().getConnection();

        SimpleDateFormat df = new SimpleDateFormat("ddMMyyyyHHmmss");
        Date date = new Date();
        String newDate = df.format(date);
        File fileOut = new File("E://" + "export" + newDate + ".xls");
        try {
             WritableWorkbook workbook1 = Workbook.createWorkbook(fileOut);
            WritableSheet sheet1 = workbook1.createSheet("First Sheet", 0);
            WritableCellFormat format = new WritableCellFormat();
            format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
            format.setBackground(Colour.GRAY_25);
            format.setVerticalAlignment(VerticalAlignment.CENTRE);
            WritableCellFormat formatRow = new WritableCellFormat();
            formatRow.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
            formatRow.setBackground(Colour.WHITE);
            formatRow.setVerticalAlignment(VerticalAlignment.CENTRE);
            
            this.dynamicQuery = new DynamicQuery(conn);
//            this.dynamicQuery.showData("SELECT BUN.BU_CODE, BUN.BU_DESCRIPTION, SO1.SO_NO, SO1.APPROVE_DATE  SO_DATE, ITM.ITEM_CODE, ITM.ITEM_DESCRIPTION, SO2.ORDER_QTY QTY,  SO2.PRICE, NVL (SO2.AMT_DISC, SO2.PRICE * NVL(SO2.PCT_DISC, 0) / 100)  DISCOUNT, (SO2.PRICE - NVL (SO2.AMT_DISC, SO2.PRICE * NVL(SO2.PCT_DISC, 0)  / 100 )) * NVL(SO2.TAX, 0) / 100 TAX\n"
//                    + "FROM TCSLS112T_HRN SO2, TCSLS111T SO1, CMITM001T ITM, TCREF001T BUN\n"
//                    + "WHERE SO2.SO_ID = SO1.SO_ID\n"
//                    + "AND SO2.ITEM_ID = ITM.ITEM_ID\n"
//                    + "AND SO1.BU_ID = BUN.BU_ID");
            this.dynamicQuery.showData(query);
            int i=0;
            for (String columnName : this.dynamicQuery.getColumnNames()) {
                System.out.println(columnName);
                 
                Label column = new Label(i,0, columnName);
                column.setCellFormat(format);
                sheet1.addCell(column);
                i++;
            }
            for (int x = 0; x < 1; x++) {
                    sheet1.setColumnView(x, 30);
                }
           
           int b=0;
            for (Vector<Object> dataRow : this.dynamicQuery.getData()) {
                String[] dataRowArray = new String[dataRow.size()];
                 int column = 0;
                
                  int a=0;
                 
                for (Object object : dataRow) {
                      try {
                        dataRowArray[column] = object.toString();
                        System.out.println(object.toString());
                        Label row = new Label(a, b + 1, object.toString());
                        row.setCellFormat(formatRow);
                        sheet1.addCell(row);
                    } catch (Exception e) {
                        dataRowArray[column] = "";
                    }
                    a++;
                    column++;
                }
                    data.add(dataRowArray);
                    b++;
            }
            int row = 0;
            int fontPointSize = 16;
            int rowHeight = (1 * fontPointSize) * 20;
          
            
            sheet1.setRowView(row, rowHeight, true);
            //   Row row = null;
//Cell cell = null;
//for (int k = 0; k < 120; k++) {
//row = sheet1.createRow(k);
//for (int j = 0; j < columns.size(); j++) {
//cell = row.createCell(j);
//if (j % 2 == 0) {
//cell.setCellValue(String.valueOf(k * j));
//}
//else {
//cell.setCellValue("Hello there");
//}
//}
//}
//for (int i = 0; i < columns.size(); i++) {
//sh.autoSizeColumn(i);
//} 
            workbook1.write();
            workbook1.close();
        } catch (SQLException | ClassNotFoundException ex) {
        } catch (IOException | WriteException ex) {
            Logger.getLogger(exequery.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
  
  
//  void fillData(JTable tabel, File file, Integer col) {
//        try {
//            WritableWorkbook workbook1 = Workbook.createWorkbook(file);
//            WritableSheet sheet1 = workbook1.createSheet("First Sheet", 0);
//            WritableCellFormat format = new WritableCellFormat();
//            format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
//            format.setBackground(Colour.GRAY_25);
//            format.setVerticalAlignment(VerticalAlignment.CENTRE);
//            WritableCellFormat formatRow = new WritableCellFormat();
//            formatRow.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
//            formatRow.setBackground(Colour.WHITE);
//            formatRow.setVerticalAlignment(VerticalAlignment.CENTRE);
//
//            TableModel model = tabel.getModel();
//            //for(int j=0;j<modelListTransact.getRowCount();j++){
//            for (int i = 0; i < col; i++) {
//
//                Label column = new Label(i, 0, model.getColumnName(i));
//                column.setCellFormat(format);
//
//                sheet1.addCell(column);
//            }
//            //}
//            int j = 0;
//            for (int i = 0; i < model.getRowCount(); i++) {
//                for (j = 0; j < col; j++) {
//                    Object row2 = model.getValueAt(i, j);
//                    String str = (row2 == null ? "" : row2.toString());
//                    Label row = new Label(j, i + 1, str);
//                    //INI untuk auto size cellnya
//                    for (int x = 0; x < model.getColumnCount(); x++) {
//                        sheet1.setColumnView(x, 30);
//                    }
//                    row.setCellFormat(formatRow);
//                    sheet1.addCell(row);
//                }
//            }
//            int row = 0;
//            int fontPointSize = 16;
//            int rowHeight = (int) ((1.5d * fontPointSize) * 20);
//            sheet1.setRowView(row, rowHeight, false);
//
////            sheet1.setColumnView(0,20);
////            sheet1.setColumnView(1,20);
////            sheet1.setColumnView(2,30);
////            sheet1.setColumnView(3,10);
////            sheet1.setColumnView(4,50);
////            sheet1.setColumnView(5,20);
////            sheet1.setColumnView(6,20);
////            sheet1.setColumnView(7,20);
//            workbook1.write();
//            workbook1.close();
//        } catch (Exception ex) {
//            ex.printStackTrace();
//        }
//    }
//  
//   public void fillData2(JTable tabel, File file,Integer col,Integer start) {
//
//        try {
//            WritableWorkbook workbook1 = Workbook.createWorkbook(file);
//            WritableSheet sheet1 = workbook1.createSheet("First Sheet", 0);  
//            WritableCellFormat format = new WritableCellFormat();
//            format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
//            format.setBackground(Colour.GRAY_25);
//            format.setVerticalAlignment(VerticalAlignment.CENTRE);
//            WritableCellFormat formatRow = new WritableCellFormat();
//            formatRow.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.getStyle(1));
//            formatRow.setBackground(Colour.WHITE);
//            formatRow.setVerticalAlignment(VerticalAlignment.CENTRE);
//
//            TableModel model =tabel.getModel();
//            //for(int j=0;j<modelListTransact.getRowCount();j++){
//            for (int i = start; i < col; i++) {
//           
//                Label column = new Label(i,0, model.getColumnName(i));
//                column.setCellFormat(format);
//                
//                sheet1.addCell(column);
//           }
//            //}
//            int j = start;
//            for (int i = 0; i < model.getRowCount(); i++) {
//                for (j = start; j < col; j++) {
//                    Object row2=model.getValueAt(i,j);
//                    String str = (row2 == null ? "" : row2.toString());  
//                    //INI untuk memulai write ke file dari kolom ke berapa dan row ke berapa
//                    Label row = new Label(j, i +1,str); 
//                    //INI untuk auto size cellnya
//                    for(int x=0;x<model.getColumnCount();x++){
//                    sheet1.setColumnView(x,30);
//                    }
//                    row.setCellFormat(formatRow);
//                    sheet1.addCell(row);
//                }
//            }
//            int row = 0;
//            int fontPointSize = 16;
//            int rowHeight = (int) ((1.5d * fontPointSize) * 20);
//            sheet1.setRowView(row, rowHeight, false);  
////            sheet1.setColumnView(0,20);
////            sheet1.setColumnView(1,20);
////            sheet1.setColumnView(2,30);
////            sheet1.setColumnView(3,10);
////            sheet1.setColumnView(4,50);
////            sheet1.setColumnView(5,20);
////            sheet1.setColumnView(6,20);
////            sheet1.setColumnView(7,20);           
//            workbook1.write();
//            workbook1.close();
//        } catch (Exception ex) {
//            ex.printStackTrace();
//        }
//    }

    public WritableSheet getSheet1() {
        return sheet1;
    }

    public void setSheet1(WritableSheet sheet1) {
        this.sheet1 = sheet1;
    }
   

}
