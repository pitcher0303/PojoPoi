package com.pojo.poi.test.sample;

import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.style.ExcelRowStyle;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;

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
            xAxis = "B"
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
            target = Category.class,
            rowStyle = @ExcelRowStyle(
                    borderBottom = BorderStyle.MEDIUM
            )
    )
    List<Category> categories = new ArrayList<>();

    @CellMeta(
            xAxis = "H"
    )
    private String issues;
}
