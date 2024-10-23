package com.pojo.poi.core.sample;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelColor;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter @Setter
@ExcelMeta(startYAxis = 1)
public class Category implements ExcelData {
    @CellMeta(
            xAxis = "E",
            cellStyle = @ExcelCellStyle(
                    borderTop = BorderStyle.DASH_DOT,
                    borderRight = BorderStyle.DASH_DOT,
                    borderBottom = BorderStyle.DASH_DOT,
                    borderLeft = BorderStyle.DASH_DOT)
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
