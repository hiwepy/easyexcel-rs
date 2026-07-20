package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.style.StyleData} for golden write.
 */
public class StyleData {

    @ExcelProperty("字符串")
    private String string;

    @ExcelProperty("字符串1")
    private String string1;

    /**
     * @return 字符串列
     */
    public String getString() {
        return string;
    }

    /**
     * @param string 字符串列
     */
    public void setString(String string) {
        this.string = string;
    }

    /**
     * @return 字符串1列
     */
    public String getString1() {
        return string1;
    }

    /**
     * @param string1 字符串1列
     */
    public void setString1(String string1) {
        this.string1 = string1;
    }
}
