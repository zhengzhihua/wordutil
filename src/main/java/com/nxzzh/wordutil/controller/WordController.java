package com.nxzzh.wordutil.controller;

import com.nxzzh.wordutil.words.XwpfTUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

@RestController
public class WordController {


    /**
     * 导出word形式
     * @param response
     */
    @RequestMapping("/exportWord")
    public void exportWord(HttpServletResponse response){

        Map<String,Object> map = new HashMap<>();
        map.put("username", "张三");
        map.put("company","杭州xx公司" );
        map.put("date","2020-04-20" );
        map.put("dept","IT部" );
        map.put("startTime","2020-04-20 08:00:00" );
        map.put("endTime","2020-04-20 08:00:00" );
        map.put("reason", "外出办公");
        map.put("time","2020-04-22" );
        try {
            response.setContentType("application/msword");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode("测试","UTF-8" );
            //String fileName = "测试"
            response.setHeader("Content-disposition","attachment;filename="+fileName+".docx" );
/*            XWPFDocument doc = WordExportUtil.exportWord07("templates/demo.docx",map);
            doc.write(response.getOutputStream());*/
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @RequestMapping("/exportWords")
    public void wordDecx(HttpServletResponse response) throws Exception {
        Map<String, Object> params = new HashMap<String, Object>();

        XwpfTUtil xwpfTUtil = new XwpfTUtil();
        String newFilses = "D:/pdffile/" + "newFile_" + XwpfTUtil.dataStr() + ".docx";
        File file1 = new File(newFilses);
        if (!file1.exists()) {
            file1.createNewFile();
        }
        response.setContentType("application/vnd.ms-excel");
        xwpfTUtil.creatWord(params,"templates/demo.docx",file1);

    }
}
