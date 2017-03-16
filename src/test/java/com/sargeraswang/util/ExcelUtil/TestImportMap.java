/**
 * @author SargerasWang
 */
package com.sargeraswang.util.ExcelUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The <code>TestImportMap</code>	
 * 
 * @author SargerasWang
 * Created at 2014年9月21日 下午5:06:17
 */
public class TestImportMap {
  @SuppressWarnings("rawtypes")
  public static void main(String[] args) throws FileNotFoundException {
    File f=new File("/Users/zach/Workspace/excel/123.xls");
    InputStream inputStream= new FileInputStream(f);
    
    ExcelLogs logs =new ExcelLogs();
    Collection<Map> importExcel = ExcelUtil.importExcel(Map.class, inputStream, "yyyy/MM/dd HH:mm:ss", logs , 0);
    
    for(Map m : importExcel){
      //检查每一行
      String line = m.get("infos").toString();

      String regex = "useTarget.{0,5}1|useTarget.{0,5}0";
      Pattern pattern = Pattern.compile(regex);
      Matcher matcher = pattern.matcher(line);


      int count = 0;
      while (matcher.find()) {
        String matchedString = matcher.group();
        System.out.print(matchedString);
        count++;
      }

      System.out.println(count);

    }
  }
}
