package com.pojo.poi.test.dto.report;

import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@ExcelMeta
public class SalesCategory implements ExcelData {
    @CellMeta(xAxis = "B")
    private String gubun;
    @CellMeta(xAxis = "C")
    private int month1;
    @CellMeta(xAxis = "D")
    private int month2;
    @CellMeta(xAxis = "E")
    private int month3;
    @CellMeta(xAxis = "F")
    private int month4;
    @CellMeta(xAxis = "G")
    private int month5;
    @CellMeta(xAxis = "H")
    private int month6;
    @CellMeta(xAxis = "I")
    private int total;
}
