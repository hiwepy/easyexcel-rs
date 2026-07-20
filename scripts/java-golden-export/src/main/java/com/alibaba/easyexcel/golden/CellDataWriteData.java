package com.alibaba.easyexcel.golden;

import java.util.Date;

import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.metadata.data.WriteCellData;

/**
 * Mirrors {@code com.alibaba.easyexcel.test.core.celldata.CellDataWriteData} for golden write.
 */
public class CellDataWriteData {

    @DateTimeFormat("yyyy年MM月dd日")
    private WriteCellData<Date> date;

    private WriteCellData<Integer> integer1;

    private Integer integer2;

    private WriteCellData<?> formulaValue;

    /**
     * @return date cell
     */
    public WriteCellData<Date> getDate() {
        return date;
    }

    /**
     * @param date date cell
     */
    public void setDate(WriteCellData<Date> date) {
        this.date = date;
    }

    /**
     * @return number WriteCellData
     */
    public WriteCellData<Integer> getInteger1() {
        return integer1;
    }

    /**
     * @param integer1 number WriteCellData
     */
    public void setInteger1(WriteCellData<Integer> integer1) {
        this.integer1 = integer1;
    }

    /**
     * @return plain integer
     */
    public Integer getInteger2() {
        return integer2;
    }

    /**
     * @param integer2 plain integer
     */
    public void setInteger2(Integer integer2) {
        this.integer2 = integer2;
    }

    /**
     * @return formula cell
     */
    public WriteCellData<?> getFormulaValue() {
        return formulaValue;
    }

    /**
     * @param formulaValue formula cell
     */
    public void setFormulaValue(WriteCellData<?> formulaValue) {
        this.formulaValue = formulaValue;
    }
}
