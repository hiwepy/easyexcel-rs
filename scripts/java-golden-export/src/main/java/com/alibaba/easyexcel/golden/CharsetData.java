package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.charset.CharsetData} for CSV charset golden write.
 */
public class CharsetData {

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("年纪")
    private Integer age;

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

    /**
     * @return 年纪列
     */
    public Integer getAge() {
        return age;
    }

    /**
     * @param age 年纪列
     */
    public void setAge(Integer age) {
        this.age = age;
    }
}
