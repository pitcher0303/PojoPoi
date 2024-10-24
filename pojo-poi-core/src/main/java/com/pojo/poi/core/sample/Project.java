package com.pojo.poi.core.sample;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelColor;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
@ExcelMeta(startYAxis = 1)
public class Project implements ExcelData {
    @CellMeta(
            xAxis = "A"
    )
    private String projectType;
    @CellMeta(
            xAxis = "B",
            cellStyle = @ExcelCellStyle(
                    borderTop = BorderStyle.DASH_DOT,
                    borderRight = BorderStyle.DASH_DOT,
                    borderBottom = BorderStyle.DASH_DOT,
                    borderLeft = BorderStyle.DASH_DOT)
    )
    private String projectName;
    @CellMeta(
            xAxis = "C"
    )
    private String projectManager;
    @CellMeta(
            xAxis = "D"
    )
    private String progressRate;

    @RowMeta(
            rowType = RowMeta.RowType.Y_RANDOM,
            target = Category.class
    )
    List<Category> categories = new ArrayList<>();

    @CellMeta(
            xAxis = "H"
    )
    private String issues;
}
