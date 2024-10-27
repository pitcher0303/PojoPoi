package com.pojo.poi.test.sample;

import com.pojo.poi.core.excel.ExcelData;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.GroupByMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.annotation.ValueMeta;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelFont;
import com.pojo.poi.core.excel.style.ExcelRowStyle;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
@ExcelMeta(
        startYAxis = 1,
        headerMetas = {
                @ValueMeta(xAxis = "A", yAxis = {1, 2}, value = "프로젝트 구분",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "B", yAxis = {1, 2}, value = "프로젝트 명",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "C", yAxis = {1, 2}, value = "담당",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "D", yAxis = {1, 2}, value = "진도율(상태)",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "E", yAxis = {1, 2}, value = "구분",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = {"F", "G", "H"}, yAxis = 1, value = "주간 수행 내역",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.THIN,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "F", yAxis = 2, value = "금주",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "G", yAxis = 2, value = "차주",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
                @ValueMeta(xAxis = "H", yAxis = 2, value = "이슈 사항",
                        cellStyle = @ExcelCellStyle(
                                borderBottom = BorderStyle.MEDIUM,
                                borderRight = BorderStyle.THIN,
                                horizontalAlignment = HorizontalAlignment.CENTER,
                                font = @ExcelFont(bold = true))
                ),
        }
)
public class Report implements ExcelData {

    @RowMeta(
            rowType = RowMeta.RowType.Y_RANDOM,
            target = Project.class,
            rowStyle = @ExcelRowStyle(
                    borderTop = BorderStyle.MEDIUM,
                    borderRight = BorderStyle.MEDIUM,
                    borderBottom = BorderStyle.MEDIUM,
                    borderLeft = BorderStyle.MEDIUM
            ),
            startYAxis = 4,
            groupBys = {
                    @GroupByMeta(xAxis = "A", dataType = GroupByMeta.DataType.CELL_DATA,
                            cellStyle = @ExcelCellStyle(
                                    wrapText = true,
                                    horizontalAlignment = HorizontalAlignment.CENTER,
                                    borderTop = BorderStyle.MEDIUM,
                                    borderRight = BorderStyle.DASH_DOT,
                                    borderBottom = BorderStyle.MEDIUM,
                                    borderLeft = BorderStyle.DASH_DOT,
                                    font = @ExcelFont(bold = true))
                    ),
                    @GroupByMeta(xAxis = "B", dataType = GroupByMeta.DataType.CELL_DATA,
                            cellStyle = @ExcelCellStyle(
                                    wrapText = true,
                                    horizontalAlignment = HorizontalAlignment.CENTER,
                                    borderTop = BorderStyle.MEDIUM,
                                    borderRight = BorderStyle.DASH_DOT,
                                    borderBottom = BorderStyle.MEDIUM,
                                    borderLeft = BorderStyle.DASH_DOT,
                                    font = @ExcelFont(bold = true))
                    ),
                    @GroupByMeta(xAxis = "C", dataType = GroupByMeta.DataType.CELL_DATA,
                            cellStyle = @ExcelCellStyle(
                                    wrapText = true,
                                    horizontalAlignment = HorizontalAlignment.CENTER,
                                    borderTop = BorderStyle.MEDIUM,
                                    borderRight = BorderStyle.DASH_DOT,
                                    borderBottom = BorderStyle.MEDIUM,
                                    borderLeft = BorderStyle.DASH_DOT)
                    ),
                    @GroupByMeta(xAxis = "D", dataType = GroupByMeta.DataType.CELL_DATA,
                            cellStyle = @ExcelCellStyle(
                                    wrapText = true,
                                    horizontalAlignment = HorizontalAlignment.CENTER,
                                    borderTop = BorderStyle.MEDIUM,
                                    borderRight = BorderStyle.DASH_DOT,
                                    borderBottom = BorderStyle.MEDIUM,
                                    borderLeft = BorderStyle.DASH_DOT,
                                    font = @ExcelFont(bold = true))
                    ),
                    @GroupByMeta(xAxis = "H", dataType = GroupByMeta.DataType.CELL_DATA,
                            cellStyle = @ExcelCellStyle(
                                    wrapText = true,
                                    verticalAlignment = VerticalAlignment.TOP,
                                    horizontalAlignment = HorizontalAlignment.LEFT,
                                    borderTop = BorderStyle.MEDIUM,
                                    borderRight = BorderStyle.DASH_DOT,
                                    borderBottom = BorderStyle.MEDIUM,
                                    borderLeft = BorderStyle.DASH_DOT)
                    ),
            }
    )
    List<Project> projectList = new ArrayList<>();
}
