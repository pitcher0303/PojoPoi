package com.pojo.poi.test.sample;

import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelFont;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@Getter
@Setter
@ExcelMeta(startYAxis = 1)
public class Category implements ExcelData {
    @CellMeta(
            xAxis = "E",
            cellStyle = @ExcelCellStyle(
                    wrapText = true,
                    horizontalAlignment = HorizontalAlignment.CENTER,
                    borderTop = BorderStyle.DASH_DOT,
                    borderRight = BorderStyle.DASH_DOT,
                    borderBottom = BorderStyle.DASH_DOT,
                    borderLeft = BorderStyle.DASH_DOT,
                    font = @ExcelFont(bold = true))
    )
    private String categoryType;
    @CellMeta(
            xAxis = "F",
            cellStyle = @ExcelCellStyle(
                    wrapText = true,
                    verticalAlignment = VerticalAlignment.TOP,
                    horizontalAlignment = HorizontalAlignment.LEFT,
                    borderTop = BorderStyle.DASH_DOT,
                    borderRight = BorderStyle.DASH_DOT,
                    borderBottom = BorderStyle.DASH_DOT,
                    borderLeft = BorderStyle.DASH_DOT)
    )
    private String thisWeek;
    @CellMeta(
            xAxis = "G",
            cellStyle = @ExcelCellStyle(
                    wrapText = true,
                    verticalAlignment = VerticalAlignment.TOP,
                    horizontalAlignment = HorizontalAlignment.LEFT,
                    borderTop = BorderStyle.DASH_DOT,
                    borderRight = BorderStyle.DASH_DOT,
                    borderBottom = BorderStyle.DASH_DOT,
                    borderLeft = BorderStyle.DASH_DOT)
    )
    private String nextWeek;
}
