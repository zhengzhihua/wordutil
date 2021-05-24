package com.nxzzh.wordutil.demo;

import com.nxzzh.wordutil.pacg.WordUtil;
import com.nxzzh.wordutil.words.XwpfTUtil;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Version 1.0.0
 * @Description
 */
public class WordDemo {

    public static void main(String[] args) {
     /*   StringBuffer messStr=new StringBuffer();
        messStr.append("    HWPFDocument是当前Word文档的代表，它的功能比WordExtractor要强。通过它我们可以读取文档中的表格、列表等，还可以对文档的内容进行新增、修改和删除操作。只是在进行完这些新增、修改和删除后相关信息是保存在HWPFDocument中的，");
        messStr.append("也就是说我们改变的是HWPFDocument，而不是磁盘上的文件。如果要使这些修改生效的话，我们可以调用HWPFDocument的write方法把修改后的HWPFDocument输出到指定的输出流中。这可以是原文件的输出流，也可以是新文件的输出流（相当于另存为）或其它输出流。");
        Map<String,Object> map = new HashMap<>();
        map.put("username", "张三");
        map.put("tbmessage", messStr.toString());
        map.put("code", "123456");
        map.put("company","xx公司" );
        map.put("date","2020-04-20" );
        map.put("dept","IT部" );
        map.put("startTime","2020-04-20 08:00:00" );
        map.put("endTime","2020-04-20 08:00:00" );
        map.put("reason", "外出办公");
        map.put("time","2020-04-22" );
        WordUtil.exportWord("templates/demo.docx","D:/" ,"生成文件1.docx" ,map );*/

        SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Map<String, Object> params = new HashMap<String, Object>();
        Map<String, Object> wdata = new HashMap<String, Object>();
        wdata.put("sch","北京一中");
        wdata.put("wd1","学生信息表，主要有学生年级、班级、家庭成员等");
        wdata.put("wd2","教师信息表，主要有教师类别、名称、等级、获得奖项等");
        wdata.put("end",sdf.format(new Date()));
        List list=new ArrayList<Object>();
        Object[] tdata1=new Object[]{"xiaozhang","男","14","初二3班","篮球","表现良好"};
        Object[] tdata2=new Object[]{"小红","女","13","初二3班","画画","表现优异"};
        list.add(tdata1);
        list.add(tdata2);
        List lists=new ArrayList<Object>();
        Object[] tdat1=new Object[]{"王老师","女","语文","A","2","高级教师"};
        Object[] tdat2=new Object[]{"张老师","女","英语","B","1","优秀教师"};
        lists.add(tdat1);
        lists.add(tdat2);
        Map<String,Object> map=new HashMap<String,Object>();
        map.put("tab0",list);
        map.put("tab1",lists);
        params.put("fdata",wdata);
        params.put("tdata",map);
        XwpfTUtil xwpfTUtil = new XwpfTUtil();
        String newFilses = "D:/pdffile/" + "newFile_" + XwpfTUtil.dataStr() + ".docx";
        File file1 = new File(newFilses);
        if (!file1.exists()) {
            try {
                file1.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        String cc=System.getProperty("user.dir");
        xwpfTUtil.creatWord(params,cc+"\\src\\main\\resources\\templates\\demos.docx",file1);

    }

}
