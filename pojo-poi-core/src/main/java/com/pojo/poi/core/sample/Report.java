package com.pojo.poi.core.sample;

import com.pojo.poi.core.excel.annotation.GroupByMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.annotation.ValueMeta;
import com.pojo.poi.core.excel.model.ExcelData;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@Getter @Setter
@ExcelMeta(
        startYAxis = 1,
        headerMetas = {
                @ValueMeta(xAxis = "A", yAxis = {1, 2}, value = "프로젝트 구분"),
                @ValueMeta(xAxis = "B", yAxis = {1, 2}, value = "프로젝트 명"),
                @ValueMeta(xAxis = "C", yAxis = {1, 2}, value = "담당"),
                @ValueMeta(xAxis = "D", yAxis = {1, 2}, value = "진도율(상태)"),
                @ValueMeta(xAxis = "E", yAxis = {1, 2}, value = "구분"),
                @ValueMeta(xAxis = {"F", "G", "H"}, yAxis = 1, value = "주간 수행 내역"),
                @ValueMeta(xAxis = "F", yAxis = 2, value = "금주"),
                @ValueMeta(xAxis = "G", yAxis = 2, value = "차주"),
                @ValueMeta(xAxis = "H", yAxis = 2, value = "이슈 사항"),
        }
)
public class Report implements ExcelData {

    @RowMeta(
            rowType = RowMeta.RowType.Y_RANDOM,
            target = Project.class,
            startYAxis = 4,
            groupBys = {
                    @GroupByMeta(xAxis = "A", dataType = GroupByMeta.DataType.CELL_DATA),
                    @GroupByMeta(xAxis = "B", dataType = GroupByMeta.DataType.CELL_DATA),
                    @GroupByMeta(xAxis = "C", dataType = GroupByMeta.DataType.CELL_DATA),
                    @GroupByMeta(xAxis = "D", dataType = GroupByMeta.DataType.CELL_DATA),
                    @GroupByMeta(xAxis = "H", dataType = GroupByMeta.DataType.CELL_DATA),
            }
    )
    List<Project> projectList = new ArrayList<>();
}
