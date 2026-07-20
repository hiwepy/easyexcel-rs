package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.head.NoHeadData} for needHead(false) golden write.
 */
public class NoHeadData {

    @ExcelProperty("字符串")
    private String string;

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
}
