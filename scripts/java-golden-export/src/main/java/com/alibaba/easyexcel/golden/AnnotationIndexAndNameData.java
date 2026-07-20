package com.alibaba.easyexcel.golden;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.annotation.AnnotationIndexAndNameData}
 * for golden write (index + name column layout).
 */
public class AnnotationIndexAndNameData {

    @ExcelProperty(value = "第四个", index = 4)
    private String index4;
    @ExcelProperty(value = "第二个")
    private String index2;
    @ExcelProperty(index = 0)
    private String index0;
    @ExcelProperty(value = "第一个", index = 1)
    private String index1;

    /**
     * @return index4 column
     */
    public String getIndex4() {
        return index4;
    }

    /**
     * @param index4 index4 column
     */
    public void setIndex4(String index4) {
        this.index4 = index4;
    }

    /**
     * @return index2 column
     */
    public String getIndex2() {
        return index2;
    }

    /**
     * @param index2 index2 column
     */
    public void setIndex2(String index2) {
        this.index2 = index2;
    }

    /**
     * @return index0 column
     */
    public String getIndex0() {
        return index0;
    }

    /**
     * @param index0 index0 column
     */
    public void setIndex0(String index0) {
        this.index0 = index0;
    }

    /**
     * @return index1 column
     */
    public String getIndex1() {
        return index1;
    }

    /**
     * @param index1 index1 column
     */
    public void setIndex1(String index1) {
        this.index1 = index1;
    }
}
