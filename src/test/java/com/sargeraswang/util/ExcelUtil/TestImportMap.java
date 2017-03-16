/**
 * @author SargerasWang
 */
package com.sargeraswang.util.ExcelUtil;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.map.HashedMap;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The <code>TestImportMap</code>
 *
 * @author SargerasWang
 *         Created at 2014年9月21日 下午5:06:17
 */
public class TestImportMap {
    @SuppressWarnings("rawtypes")
    public static void main(String[] args) throws Exception {
        File f = new File("/Users/zach/Workspace/excel/123.xls");
        InputStream inputStream = new FileInputStream(f);

        ExcelLogs logs = new ExcelLogs();
        Collection<Map> importExcel = ExcelUtil.importExcel(Map.class, inputStream, "yyyy/MM/dd HH:mm:ss", logs, 0);

        //检查成人数
        for (Map m : importExcel) {
            //检查每一行
            String line = m.get("infos").toString();

            String regex = "useTarget.{0,5}1";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(line);


            int count = 0;
            while (matcher.find()) {
                String matchedString = matcher.group();
                System.out.print(matchedString);
                count++;
            }

            String resString = m.get("resIds").toString();
            //查找资源
            if (null != resString) {
                String[] reses = resString.split(",");
                m.put("resNum", reses.length);
            }else{
                m.put("resNum", "0");
            }

            m.put("adultNum", count);
        }

        //检查儿童数
        for (Map m : importExcel) {
            //检查每一行
            String line = m.get("infos").toString();

            String regex = "useTarget.{0,5}2";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(line);


            int count = 0;
            while (matcher.find()) {
                String matchedString = matcher.group();
                System.out.print(matchedString);
                count++;
            }

            m.put("childNum", count);
        }

        //检查其他3-11
        for (Map m : importExcel) {
            //检查每一行
            String line = m.get("infos").toString();

            String regex = "useTarget.{0,5}3|useTarget.{0,5}4|useTarget.{0,5}5|useTarget.{0,5}6|useTarget.{0,5}7|useTarget.{0,5}8|useTarget.{0,5}9|useTarget.{0,5}10|useTarget.{0,5}11";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(line);


            int count = 0;
            while (matcher.find()) {
                String matchedString = matcher.group();
                System.out.print(matchedString);
                count++;
            }

            m.put("otherNum", count);
        }

        //检查0
        for (Map m : importExcel) {
            //检查每一行
            String line = m.get("infos").toString();

            String regex = "useTarget.{0,5}0";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(line);


            int count = 0;
            while (matcher.find()) {
                String matchedString = matcher.group();
                System.out.print(matchedString);
                count++;
            }

            m.put("zeroNum", count);
        }

        List<Map<String, Object>> result = new ArrayList<>();
        if (CollectionUtils.isNotEmpty(importExcel)) {
            int count = 0;
            for (Map m : importExcel) {
                count++;
                Map<String, Object> tMap = new HashMap();
                tMap.put("number", count);
                tMap.put("baseId", m.get("baseId"));
                tMap.put("resIds", m.get("resIds"));
                tMap.put("adultNum", m.get("adultNum"));
                tMap.put("childNum", m.get("childNum"));
                tMap.put("otherNum", m.get("otherNum"));
                tMap.put("zeroNum", m.get("zeroNum"));
                tMap.put("resNum", m.get("resNum"));
                tMap.put("infos", m.get("infos"));

                result.add(tMap);
            }
        }

        File outFile = new File("/Users/zach/Workspace/excel/123Result.xls");
        OutputStream out = new FileOutputStream(outFile);
        ExcelUtil.exportExcel(new String[]{"number", "baseId", "resIds", "adultNum", "childNum", "otherNum", "zeroNum", "resNum", "infos"}, result, out);
        out.close();
    }
}
