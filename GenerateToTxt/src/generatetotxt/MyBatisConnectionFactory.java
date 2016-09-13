/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generatetotxt;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.Reader;
import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

/**
 *
 * @author azec
 */
public class MyBatisConnectionFactory {
    
    
     private static SqlSessionFactory sqlSessionFactory;

  public static SqlSessionFactory getSqlSessionFactory()
  {
    return sqlSessionFactory;
  }

  static
  {
    try
    {
      String resource = "generatetotxt/UtilMapper.xml";
      Reader reader = Resources.getResourceAsReader(resource);

      if (sqlSessionFactory == null)
      {
        sqlSessionFactory = new SqlSessionFactoryBuilder().build(reader);
      }

    }
    catch (FileNotFoundException fileNotFoundException)
    {
      fileNotFoundException.printStackTrace();
    }
    catch (IOException iOException)
    {
      iOException.printStackTrace();
    }
  }

}
