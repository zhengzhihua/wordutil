package com.nxzzh.wordutil.pacg;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.Assert;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

/**
 * @Version 1.0.0
 * @Description
 */
public class WordUtil {


    /**
     * 生成word
     * @param templatePath
     * @param temDir
     * @param fileName
     * @param params
     */
    public static void exportWord(String templatePath, String temDir, String fileName, Map<String,Object> params){
        Assert.notNull(templatePath, "模板路径不能为空");
        Assert.notNull(temDir, "临时文件路径不能为空");
        Assert.notNull(fileName, "导出文件名不能为空");
        Assert.isTrue(fileName.endsWith(".docx"), "word导出请使用docx格式");
        if (!temDir.endsWith("/")) {
            temDir = temDir + File.separator;
        }
        File dir = new File(temDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
       /* try {
            XWPFDocument doc = WordExportUtil.exportWord07(templatePath, params);

            String tmpPath = temDir + fileName;
            FileOutputStream fos = new FileOutputStream(tmpPath);
            doc.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }*/
    }

}


