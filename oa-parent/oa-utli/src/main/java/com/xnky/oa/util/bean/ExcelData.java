package com.xnky.oa.util.bean;

import java.util.List;
import java.util.Map;
/**
 * 
 * 描述:为导出Excel文件封装数据
 * 封装Excel文件数据注意事项:
 * 1.heads(标题行):是一个字符串, 标题与标题之间以逗号分隔如:hears="学号,姓名,性别"
 * 2.data: 表示:主数据集合,data是个字符串集合,包含了所有主数据,集合中的一个字符串元素就是Excel表格中的一行记录;
 * 列数据项与列数项之间用逗号分隔, 如果某个列数据项值为空, 就用null代替,如: data.add("001,老李,null,15056786789,null");
 * 3.assoceData 关联数据集合,主数据集合的字符串元素的下标在关联数据集合就是对应的关联数据.是个双list集合字符串,包含了所有主数据的关联数据集合,
 * 它的每一个元素就是某一行主数据的关联数据集合
 * 4.widthAndHeiht map的key 表示列的位置,map的value 表示列的宽度, 如:widthAndHeiht.put(3, 5000); 第4列宽度为5000英寸
 * 5.length 合并单元格的个数
 * @author sunxuefeng 2017年7月28日 下午5:13:35
 * @version 1.0
 */

public class ExcelData {

    private String                excelName;    // Excel文件名称
    private String                sheetName;    // 工作表的名称
    private String                heads;        // 标题行
    private List<String>          data;         // 主数据
    private Map<Integer, Integer> widthAndHeiht;// 单元格的宽高,如:xSheet.setColumnWidth(10,1000);
                                                // 10表示:map的key, //
                                                // 10,1000表示:map的value
    private List<List<String>>    assoceData;   // 关联数据行
    private int                   length;       // 合并单元格的个数 没有就不用填



    public String getHeads() {
        return heads;
    }



    public void setHeads(String heads) {
        this.heads = heads;
    }



    public int getLength() {
        return length;
    }



    public void setLength(int length) {
        this.length = length;
    }



    public List<List<String>> getAssoceData() {
        return assoceData;
    }



    public void setAssoceData(List<List<String>> assoceData) {
        this.assoceData = assoceData;
    }



    public String getExcelName() {
        return excelName;
    }



    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }



    public String getSheetName() {
        return sheetName;
    }



    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }



    public List<String> getData() {
        return data;
    }



    public void setData(List<String> data) {
        this.data = data;
    }



    public Map<Integer, Integer> getWidthAndHeiht() {
        return widthAndHeiht;
    }



    public void setWidthAndHeiht(Map<Integer, Integer> widthAndHeiht) {
        this.widthAndHeiht = widthAndHeiht;
    }

}
