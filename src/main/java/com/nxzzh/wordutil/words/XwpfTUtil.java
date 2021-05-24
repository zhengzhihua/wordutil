package com.nxzzh.wordutil.words;

import org.apache.poi.xwpf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.SimpleFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XwpfTUtil {

    /*
     *输出文档
     */
    public void creatWord(Map<String,Object> wodata, String tempFile, File newFile){
        XWPFDocument xwpfDocument;
        File file = new File(tempFile);
        OutputStream outputStream = null;
        InputStream inputStream = null;
        try {
            inputStream=new FileInputStream(file);
            xwpfDocument=new XWPFDocument(inputStream);
            replaceInPara(xwpfDocument, (Map<String, Object>) wodata.get("fdata"));
            replaceInTable(xwpfDocument, (Map<String, Object>) wodata.get("tdata"));
            outputStream=new FileOutputStream(newFile);
            xwpfDocument.write(outputStream);
            close(inputStream);
            close(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }

    }

    public static String dataStr(){
        Date date= new Date();
        SimpleDateFormat simpleFormatter =new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
        return simpleFormatter.format(date);
    }

    /**
     * 替换段落里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    /**
     * 替换段落里面的变量
     *
     * @param para   要替换的段落
     * @param params 参数
     */
    public void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        List<XWPFRun> runs;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                System.out.print(runs.size()+"----"+params.size());
                XWPFRun run = runs.get(i);
                for(Map.Entry<String, Object> entry: params.entrySet()){
                    if(run.toString().contains("${"+entry.getKey()+"}")){
                        System.out.print(run.toString());
                        String text = run.toString().replace("${" + entry.getKey() + "}", entry.getValue().toString());
                        run.setText(text,0);
                    }
                }

            }

        /*    for (String key : params.keySet()) {
                if (str.equals(key)) {
                    para.createRun().setText((String) params.get(key));
                    break;
                }
            }*/
        }
    }

    /**
     * 替换表格里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        while (iterator.hasNext()) {
            for(int num=0; num<params.size();num++){
                table = iterator.next();
                this.insertTable(table,(List)params.get("tab"+num));
            }
        }
    }

    /*
    * 表格插入数据
    * */

    private void insertTable( XWPFTable table,List<Object[]> tabdata){
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        if(null != tabdata && tabdata.size()>0){
            for(int rowN = 1 ;rowN <table.getRows().size() ;rowN++){
                rows = table.getRows();
                for(int cellN = 0 ;cellN <rows.get(rowN).getTableCells().size() ;cellN++){
                    cells = rows.get(rowN).getTableCells();
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
