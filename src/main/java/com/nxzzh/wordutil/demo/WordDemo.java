package com.nxzzh.wordutil.demo;

import com.nxzzh.wordutil.pacg.WordUtil;

import java.util.HashMap;
import java.util.Map;

/**
 * @Version 1.0.0
 * @Description
 */
public class WordDemo {

    public static void main(String[] args) {
        StringBuffer messStr=new StringBuffer();
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

        WordUtil.exportWord("templates/demo.docx","D:/" ,"生成文件1.docx" ,map );

    }

}
