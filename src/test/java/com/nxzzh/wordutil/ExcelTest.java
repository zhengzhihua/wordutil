package com.nxzzh.wordutil;

import com.nxzzh.wordutil.excels.ExcelUtil;
import com.nxzzh.wordutil.words.XwpfTUtil;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelTest {

    @Test
    public void Test1() throws Exception {
        System.out.print("hello");
        ExcelUtil excelUtil=new ExcelUtil();
        //获取项目相对路径根目录
        String cc=System.getProperty("user.dir");
        //生成的excel文件
        String newFilses = cc+"\\src\\main\\resources\\pdffile\\" + "excFile_" + XwpfTUtil.dataStr() + ".xlsx";

        //excel模板路径
        String tempExcel=cc+"\\src\\main\\resources\\templates\\demo.xlsx";
        Map<String, Object> data=new HashMap<>();
        List<Object> head=new ArrayList<>();
        List<Object> tails=new ArrayList<>();
        Object[] rowdata1=new Object[]{"101","张三","112","120","","98","97","90","430"};
        Object[] rowdata2=new Object[]{"101","王五","111","113","104","95","70","","415"};
        tails.add(rowdata1);
        tails.add(rowdata2);
        data.put("data1",tails);
        excelUtil.generateExcel(tempExcel,newFilses,data);

    }
}
