/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generatetotxt;

import id.my.berkah.util.DynamicQuery;
import java.io.IOException;
import jxl.write.WriteException;
/**
 *
 * @author azec
 */
public class GenerateToTxt extends exeTXT{
 DynamicQuery dynamicQuery;
   
    public static void main(String[] args) throws IOException, WriteException {

     exeTXT exT= new exeTXT();
     exT.Query1("select user_id, user_name, user_description from fnd_users");
    
    }
//                this.dynamicQuery.showData("SELECT BUN.BU_CODE, BUN.BU_DESCRIPTION, SO1.SO_NO, SO1.APPROVE_DATE  SO_DATE, ITM.ITEM_CODE, ITM.ITEM_DESCRIPTION, SO2.ORDER_QTY QTY,  SO2.PRICE, NVL (SO2.AMT_DISC, SO2.PRICE * NVL(SO2.PCT_DISC, 0) / 100)  DISCOUNT, (SO2.PRICE - NVL (SO2.AMT_DISC, SO2.PRICE * NVL(SO2.PCT_DISC, 0)  / 100 )) * NVL(SO2.TAX, 0) / 100 TAX\n"
//                    + "FROM TCSLS112T_HRN SO2, TCSLS111T SO1, CMITM001T ITM, TCREF001T BUN\n"
//                    + "WHERE SO2.SO_ID = SO1.SO_ID\n"
//                    + "AND SO2.ITEM_ID = ITM.ITEM_ID\n"
//                    + "AND SO1.BU_ID = BUN.BU_ID");
    
}