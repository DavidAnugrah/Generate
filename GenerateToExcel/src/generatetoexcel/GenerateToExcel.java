/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generatetoexcel;

import java.io.IOException;
import jxl.write.WriteException;

/**
 *
 * @author azec
 */
public class GenerateToExcel extends exequery{
    
    public static void main(String[] args) throws IOException, WriteException{
   
        exequery exeQuery = new exequery();
//        exeQuery.fillData2(null, fileOut, Integer.SIZE, Integer.SIZE);
        exeQuery.Query1("select user_id, user_name, user_description from fnd_users");


        }
    }

    

