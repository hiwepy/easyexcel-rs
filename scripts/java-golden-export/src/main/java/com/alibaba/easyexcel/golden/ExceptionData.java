package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.exception.ExceptionData} for golden write.
 */
public class ExceptionData {

    @ExcelProperty("姓名")
    private String name;

    /**
     * @return 姓名列
     */
    public String getName() {
        return name;
    }

    /**
     * @param name 姓名列
     */
    public void setName(String name) {
        this.name = name;
    }
}
