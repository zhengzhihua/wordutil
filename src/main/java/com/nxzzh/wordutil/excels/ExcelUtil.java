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

    private void editXSSFWorkbook(XSSFWorkbook hssfWork,Map<String, Object> excelData){
         for (int shet=0;shet<hssfWork.getNumberOfSheets();shet++){
             XSSFSheet sheet=hssfWork.getSheetAt(shet);
             editXssfheet(sheet,excelData);
         }
    }

    private void editXssfheet(XSSFSheet sheet, Map<String, Object> excelData) {
        //    int rows=sheet.getPhysicalNumberOfRows();
        List<Object> datas = (List<Object>) excelData.get("data1");
        for (int i = 0; i < datas.size(); i++) {
            Object[] calsdata = (Object[]) datas.get(i);
            XSSFRow xssfRow = sheet.createRow(i + 2);

            //           xssfRow.createCell(((Object[])calsdata).length);
            System.out.println(((Object[]) calsdata).length);
            for (int j = 0; j < ((Object[]) calsdata).length; j++) {
                xssfRow.createCell(j);
                xssfRow.getCell(j).setCellValue(((Object[]) calsdata)[j].toString());
            }

        }
    }

}
