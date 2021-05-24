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

    /*String filePath = "/sta.docx";
    InputStream is;
    XWPFDocument doc;
    Map<String, Object> params = new HashMap<String, Object>();
    {
        params.put("${name}", "xxx");
        params.put("${sex}", "男");
        params.put("${political}", "共青团员");
        params.put("${place}", "sssss");
        params.put("${classes}", "3102");
        params.put("${id}", "213123123");
        params.put("${qq}", "213123");
        params.put("${tel}", "312313213");
        params.put("${oldJob}", "sadasd");
        params.put("${swap}", "是");
        params.put("${first}", "asdasd");
        params.put("${second}", "综合事务部");
        params.put("${award}", "asda");
        params.put("${achievement}", "完成科协网站的开发");
        params.put("${advice}", "没有建议");
        params.put("${attach}", "无");
        try {
            is = new FileInputStream(filePath);
            doc = new XWPFDocument(is);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/


    /**
     * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
     *
     * @throws Exception
     */
    /*@Test
    public void testTemplateWrite() throws Exception {
        //替换段落里面的变量
        this.replaceInPara(doc, params);
        //替换表格里面的变量
        this.replaceInTable(doc, params);
        OutputStream os = new FileOutputStream("D:\\sta1.docx");
        doc.write(os);
        this.close(os);
        this.close(is);
    }*/

    /*@Test
    public void myTest1() throws Exception {
        *//*Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            List<XWPFRun> runs = para.getRuns();
            para.removeRun(0);
            para.insertNewRun(0).setText("hello");
        }
        OutputStream os = new FileOutputStream("D:\\sta1.docx");
        doc.write(os);
        this.close(os);
        this.close(is);*//*
        System.out.println(this.matcher("报告日期：${reportDate}").find());
    }*/

    /*@Test
    public void myReplaceInPara() {
//        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
//        XWPFParagraph para;
//        while (iterator.hasNext()) {
//            para = iterator.next();
//            List<XWPFRun> runs = para.getRuns();
//
//
//        }
        System.out.println('{'=='{');
    }*/


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
            inputStream.close();
            outputStream.close();
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
        Matcher matcher;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                for(Map.Entry<String, Object> entry: params.entrySet()){
                    if(run.toString().contains("${"+entry.getKey()+"}")){
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
