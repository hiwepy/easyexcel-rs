package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.cache.CacheData} for golden write.
 */
public class CacheData {

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("年龄")
    private Long age;

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
     * @return 年龄列
     */
    public Long getAge() {
        return age;
    }

    /**
     * @param age 年龄列
     */
    public void setAge(Long age) {
        this.age = age;
    }
}
