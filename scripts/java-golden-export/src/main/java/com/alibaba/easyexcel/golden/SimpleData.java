package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Minimal row model mirroring {@code com.alibaba.easyexcel.test.core.simple.SimpleData}
 * for Java write → golden export.
 */
public class SimpleData {

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
