package com.pojo.poi.test.dto.report;

import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.annotation.*;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelFont;
import com.pojo.poi.core.excel.style.ExcelRowStyle;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;

import java.util.List;

@Getter
@Setter
@ExcelMeta(
        startYAxis = 2,
        headerMetas = {
                @ValueMeta(xAxis = "B", cellStyle = @ExcelCellStyle(font = @ExcelFont(bold = true)), value = "상반기 월별 매출 현황"),
                @ValueMeta(xAxis = "G", value = "년도"),
                @ValueMeta(xAxis = "I", value = "단위"),
        }
)
public class Sales implements ExcelData {
    @CellMeta(xAxis = "H")
    private int year;

    @RowMeta(
            startYAxis = 2,
            target = SalesCategory.class,
            rowStyle = @ExcelRowStyle(
                    borderTop = BorderStyle.MEDIUM,
                    borderBottom = BorderStyle.MEDIUM),
            headerMetas = {
                    @ValueMeta(xAxis = "B", value = "구분"),
                    @ValueMeta(xAxis = "C", value = "1월"),
                    @ValueMeta(xAxis = "D", value = "2월"),
                    @ValueMeta(xAxis = "E", value = "3월"),
                    @ValueMeta(xAxis = "F", value = "4월"),
                    @ValueMeta(xAxis = "G", value = "5월"),
                    @ValueMeta(xAxis = "H", value = "6월"),
                    @ValueMeta(xAxis = "I", value = "총합계"),
            },
            metaOrder = @MetaOrder(1))
    List<SalesCategory> categories;

    @RowMeta(
            target = SalesCategory.class,
            metaOrder = @MetaOrder(value = 2, type = MetaOrder.Type.Y_REFERENCES, referenceMetaOrder = 1))
    List<SalesCategory> totalCategories;
}
