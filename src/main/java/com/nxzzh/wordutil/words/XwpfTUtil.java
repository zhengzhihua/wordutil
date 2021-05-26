package com.nxzzh.wordutil.words;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XwpfTUtil {

    /**
     * @description 根据word模板创建word文档
     * @param wodata  文档数据
     * @param tempFile  模板路径
     * @param newFile  创建的文档路径
     */
    public void creatWord(Map<String,Object> wodata, String tempFile, File newFile){
        XWPFDocument xwpfDocument;
        File file = new File(tempFile);
        OutputStream outputStream = null;
        InputStream inputStream = null;
        try {
            inputStream=new FileInputStream(file);
            xwpfDocument=new XWPFDocument(inputStream);
            //获取word段落文字的数据
            replaceInPara(xwpfDocument, (Map<String, Object>) wodata.get("fdata"));
            //获取word表格的数据
            replaceInTable(xwpfDocument, (Map<String, Object>) wodata.get("tdata"));
            outputStream=new FileOutputStream(newFile);
            //将操作完的xwpfDocument写入输出流中，输出文件中
            xwpfDocument.write(outputStream);
            close(inputStream);
            close(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }

    }

    /**
     * @description  获取string类型的系统时间
     * @return  string时间
     */
    public static String dataStr(){
        Date date= new Date();
        SimpleDateFormat simpleFormatter =new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
        return simpleFormatter.format(date);
    }

    /**
     * @description 替换段落里面的变量
     * @param doc    要替换的文档
     * @param params 替换的数据
     */
    public void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;

        while (iterator.hasNext()) {
            para = iterator.next();
            para.setAlignment(ParagraphAlignment.CENTER);
            this.replaceInPara(para, params);
        }
    }

    /**
     * @description 替换段落里面的变量
     * @param para   要替换的段落
     * @param params 数据
     */
    public void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        List<XWPFRun> runs;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                System.out.print(runs.size()+"----"+params.size());
                XWPFRun run = runs.get(i);
                for(Map.Entry<String, Object> entry: params.entrySet()){
                    //判断${key}标签，通过数据替换${key}标签
                    if(run.toString().contains("${"+entry.getKey()+"}")){
                        System.out.print(run.toString());
                        String text = run.toString().replace("${" + entry.getKey() + "}", entry.getValue().toString());
                        run.setText(text,0);
                    }
                }

            }
        }
    }

    /**
     * @description 遍历文档中的表格，对指定表格指定数据
     * @param doc    要替换的表格
     * @param params 数据
     */
    public void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        while (iterator.hasNext()) {
            for(int num=0; num<params.size();num++){
                //获取word中的表格，对表格指定数据
                table = iterator.next();
                this.insertTable(table,(List)params.get("tab"+num));
            }
        }
    }

    /**
     * @description  替换指定表格中的数据
     * @param table  指定表格
     * @param tabdata  表格数据
     */

    private void insertTable( XWPFTable table,List<Object[]> tabdata){
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        if(null != tabdata && tabdata.size()>0){
            for(int rowN = 1 ;rowN <table.getRows().size() ;rowN++){
                //获取表格行
                rows = table.getRows();
                for(int cellN = 0 ;cellN <rows.get(rowN).getTableCells().size() ;cellN++){
                    //获取表格中指定行中的单元格
                    cells = rows.get(rowN).getTableCells();
                    //对指定单元格赋值
                    cells.get(cellN).setText((tabdata.get(rowN-1)[cellN]).toString());
                }
            }
        }

    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 关闭输入流
     *
     * @param is
     */
    public void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    public void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
