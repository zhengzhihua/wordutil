package com.nxzzh.wordutil.excels;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;
import java.util.Map;

/*
 *
 * 通过excel模板生成excel表格
 */
public class ExcelUtil {

    /*
     * @tempFile  :模板文件路径
     * @fileStr  :生成文件路径
     * @excelData  :生成文件中的数据
     * @decripion :通过模板文件和数据生成excel文件
     *
     */
    public void generateExcel(String tempFile, String fileStr, Map<String, Object> excelData) throws Exception {
        File file1 = new File(fileStr);
        if (!file1.exists()) {
            try {
                file1.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        OutputStream output=new FileOutputStream(file1);
        InputStream input = new FileInputStream(tempFile);
        XSSFWorkbook hssfWork=new XSSFWorkbook(input);
        int sheetNum=hssfWork.getNumberOfSheets();
        System.out.println(sheetNum);
        editXSSFWorkbook(hssfWork,excelData);
        hssfWork.write(output);
        input.close();
        output.close();
    }
    /*
     * 遍历表格文件的sheet
     */
    private void editXSSFWorkbook(XSSFWorkbook hssfWork,Map<String, Object> excelData){
         for (int shet=0;shet<hssfWork.getNumberOfSheets();shet++){        //遍历表格文件的sheet
             XSSFSheet sheet=hssfWork.getSheetAt(shet);
             editXssfheet(sheet,excelData);                                //对每个sheet进行操作
         }
    }

    /**
     * @description
     * @param sheet
     * @param excelData
     */
    private void editXssfheet(XSSFSheet sheet, Map<String, Object> excelData) {
        //    int rows=sheet.getPhysicalNumberOfRows();
        List<Object> datas = (List<Object>) excelData.get("data1");
        for (int i = 0; i < datas.size(); i++) {       //根据数据list的大小创建表格行，
            Object[] calsdata = (Object[]) datas.get(i);
            XSSFRow xssfRow = sheet.createRow(i + 2);     //从表格第几行开始生成指定行

            //           xssfRow.createCell(((Object[])calsdata).length);
            System.out.println(((Object[]) calsdata).length);
            for (int j = 0; j < ((Object[]) calsdata).length; j++) {         //通过数据中行数据的大小创建cell，行的每一格
                xssfRow.createCell(j);                                   //创建xssfRow的指定格，xssfRow是创建的行
                xssfRow.getCell(j).setCellValue(((Object[]) calsdata)[j].toString());       //向创建的行添加数据
            }

        }
    }

}
