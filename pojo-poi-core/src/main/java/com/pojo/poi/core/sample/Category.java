package com.pojo.poi.core.sample;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.ExcelData;
import lombok.Getter;
import lombok.Setter;

@Getter @Setter
@ExcelMeta(startYAxis = 1)
public class Category implements ExcelData {
    @CellMeta(
            xAxis = "E"
    )
    private String categoryType;
    @CellMeta(
            xAxis = "F"
    )
    private String thisWeek;
    @CellMeta(
            xAxis = "G"
    )
    private String nextWeek;
}
