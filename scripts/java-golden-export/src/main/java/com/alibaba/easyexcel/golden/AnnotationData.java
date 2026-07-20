package com.alibaba.easyexcel.golden;

import java.util.Date;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.annotation.AnnotationData}.
 */
@ColumnWidth(50)
@HeadRowHeight(50)
@ContentRowHeight(100)
public class AnnotationData {

    @ExcelProperty("日期")
    @DateTimeFormat("yyyy年MM月dd日HH时mm分ss秒")
    private Date date;

    @ExcelProperty("数字")
    @NumberFormat("#.##%")
    private Double number;

    @ExcelIgnore
    private String ignore;

    /**
     * @return 日期列
     */
    public Date getDate() {
        return date;
    }

    /**
     * @param date 日期列
     */
    public void setDate(Date date) {
        this.date = date;
    }

    /**
     * @return 数字列
     */
    public Double getNumber() {
        return number;
    }

    /**
     * @param number 数字列
     */
    public void setNumber(Double number) {
        this.number = number;
    }

    /**
     * @return 忽略列
     */
    public String getIgnore() {
        return ignore;
    }

    /**
     * @param ignore 忽略列
     */
    public void setIgnore(String ignore) {
        this.ignore = ignore;
    }
}
