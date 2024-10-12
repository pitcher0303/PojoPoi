package com.pojo.poi.core.sample;

import com.pojo.poi.core.excel.annotation.ModelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.annotation.ValueMeta;
import com.pojo.poi.core.excel.model.ExcelData;

import java.util.List;

@ModelMeta(
        headerMetas = {
                @ValueMeta(xAxis = "A", yAxis = {1, 2}, value = "목록"),
                @ValueMeta(xAxis = "B", yAxis = {1, 2}, value = "프로젝트"),
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
            target = Project.class,
            startYAxis = 3
    )
    List<Project> projectList;
}
