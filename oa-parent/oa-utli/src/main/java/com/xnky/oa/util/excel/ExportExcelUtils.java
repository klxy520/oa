package com.xnky.oa.util.excel;

import java.net.URLEncoder;
import java.util.List;
import java.util.Map;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.xnky.oa.util.bean.ExcelData;

/**
 * 
 * 描述：导出excel文件模板
 * 
 * @author sunxuefeng 2017年7月28日 上午11:20:48
 * @version 1.0
 */
public class ExportExcelUtils {
    private static final Logger LOG = Logger.getLogger(ExportExcelUtils.class);



    /**
     * 
     * 描述：输出Excel文件
     * 
     * @param response
     * @param excelData
     *            为导出Excel文件封装数据
     * @author sunxuefeng 2017年7月31日 上午9:58:10
     * @version 1.0
     * @throws Exception
     */
    public static void exportFile(HttpServletResponse response, ExcelData excelData) {
        ServletOutputStream outputStream = null;
        try {
            outputStream = response.getOutputStream();
            response.setHeader("Content-disposition",
                    "attachment; filename = " + URLEncoder.encode(excelData.getExcelName() + ".xlsx", "UTF-8"));
            response.setContentType("application/octet-streem");
            List<List<String>> assoceData = excelData.getAssoceData();
            XSSFWorkbook workbook = null;
            if (assoceData != null && assoceData.size() > 0) {
                workbook = setSheetContent(excelData);
            } else {
                workbook = setSheetContentNO(excelData);
            }
            workbook.write(outputStream);
        } catch (Exception e) {
            LOG.error("Excel文件导出失败:" + e.getMessage());
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }



    /**
     * 
     * 描述： set Sheet标题行
     * 
     * @param xWorkbook
     * @param xSheet
     * @param headers标题行
     * @author sunxuefeng 2017年7月28日 下午1:13:36
     * @version 1.0
     */
    private static void setSheetHeader(XSSFWorkbook xWorkbook, XSSFSheet xSheet, String headers) {
        // 设置样式
        CellStyle cs = xWorkbook.createCellStyle();
        // 设置水平垂直居中
        cs.setAlignment(CellStyle.ALIGN_CENTER);
        cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 设置字体
        Font headerFont = xWorkbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        headerFont.setFontName("宋体");
        cs.setFont(headerFont);
        cs.setWrapText(false);// 是否自动换行
        XSSFRow xRow0 = xSheet.createRow(0);
        String[] header = headers.split(",");
        for (int i = 0, length = header.length; i < length; i++) {
            XSSFCell xCell = xRow0.createCell(i);
            xCell.setCellStyle(cs);
            xCell.setCellValue(header[i]);
        }
    }



    /***
     * 
     * 描述： 设置Sheet页内容(有关联数据)
     * 
     * @param excelData
     * @return
     * @author sunxuefeng 2017年7月31日 上午9:59:24
     * @version 1.0
     */
    private static XSSFWorkbook setSheetContent(ExcelData excelData) {
        XSSFWorkbook xWorkbook = new XSSFWorkbook();
        XSSFSheet xSheet = xWorkbook.createSheet(excelData.getSheetName());
        Map<Integer, Integer> map = excelData.getWidthAndHeiht();
        CellStyle cs = xWorkbook.createCellStyle(); // 设置样式
        cs.setAlignment(CellStyle.ALIGN_CENTER);
        cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cs.setWrapText(false); // 是否自动换行
        setSheetHeader(xWorkbook, xSheet, excelData.getHeads());
        if (map != null && map.size() > 0) { // 设置单元格的宽高
            for (Integer key : map.keySet()) {
                xSheet.setColumnWidth(key, map.get(key));
            }
        }
        List<String> data = excelData.getData();
        List<List<String>> assoceDatas = excelData.getAssoceData();
        if (data != null && data.size() > 0) {
            int l = 0, ii = 0;
            for (int i = 0; i < data.size(); i++) {
                List<String> assoceData = assoceDatas.get(i);
                int assoceDataSize = 0;
                boolean assoceDataFlag = false;
                if (assoceData != null && assoceData.size() > 0) {
                    assoceDataSize = assoceData.size();
                    assoceDataFlag = true;
                    if (i == 0) {
                        ii = 0;
                        l = assoceDataSize;
                    } else {
                        ii = l;
                        l += assoceDataSize;
                    }
                } else {
                    if (i == 0) {
                        ii = 0;
                        l = 1;
                    } else {
                        ii = l;
                        l += 1;
                    }
                }
                for (int k = ii, m = 0, lengths = l; k < lengths; k++, m++) {// 控制行行:
                    XSSFRow xRow = xSheet.createRow(k + 1);
                    int j = 0;
                    j = setValue(j, 0, data.get(i), xRow, cs);// 设置主数据
                    int length = data.get(i).split(",").length;
                    if (assoceDataFlag) {
                        setValue(j, length, assoceData.get(m), xRow, cs);// 设置关联数据
                    }
                }
                if (assoceDatas.size() > 1) { // 合并单元格
                    for (int m = 0; m < excelData.getLength(); m++) {
                        if (assoceDataFlag) {
                            int endRows = xSheet.getLastRowNum();
                            int startRow = endRows - assoceDataSize + 1;
                            xSheet.addMergedRegion(new CellRangeAddress(startRow, endRows, m, m));
                        }
                    }
                }
            }
        }
        return xWorkbook;
    }



    /***
     * 
     * 描述： 设置Sheet页内容(无关联数据)
     * 
     * @param excelData
     * @return
     * @author sunxuefeng 2017年7月31日 上午9:59:24
     * @version 1.0
     */
    private static XSSFWorkbook setSheetContentNO(ExcelData excelData) throws Exception {
        XSSFWorkbook xWorkbook = new XSSFWorkbook();
        XSSFSheet xSheet = xWorkbook.createSheet(excelData.getSheetName());
        Map<Integer, Integer> map = excelData.getWidthAndHeiht();
        CellStyle cs = xWorkbook.createCellStyle(); // 设置样式
        cs.setAlignment(CellStyle.ALIGN_CENTER);
        cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cs.setWrapText(false); // 是否自动换行
        setSheetHeader(xWorkbook, xSheet, excelData.getHeads());
        if (map != null && map.size() > 0) { // 设置单元格的宽
            for (Integer key : map.keySet()) {
                xSheet.setColumnWidth(key, map.get(key));
            }
        }
        List<String> data = excelData.getData();
        if (data != null && data.size() > 0) {
            for (int i = 0; i < data.size(); i++) {
                XSSFRow xRow = xSheet.createRow(i + 1);
                int j = 0;
                j = setValue(j, 0, data.get(i), xRow, cs);// 设置主数据
            }
        }
        return xWorkbook;
    }



    /**
     * 
     * 描述：为单元格设置值
     * 
     * @author sunxuefeng 2017年8月1日 上午9:45:36
     * @version 1.0
     */
    private static int setValue(int j, int i, String data, XSSFRow xRow, CellStyle cs) {
        String[] strings = data.split(",");
        for (; j < strings.length + i; j++) {
            String dString = strings[j - i];
            if (dString != null && !dString.equals("") && !dString.equals("null")) {
                XSSFCell xCell = xRow.createCell(j);
                xCell.setCellStyle(cs);
                xCell.setCellValue(dString);
            }
        }
        return j;
    }
}
