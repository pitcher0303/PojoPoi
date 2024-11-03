package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

public class ExcelReader {
    public static <T extends ExcelData> T readSheet(Class<T> to, Sheet sheet) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        T toInstnace = to.getConstructor().newInstance();
        if (!to.isAnnotationPresent(ExcelMeta.class)) return toInstnace;

        ExcelMeta excelMeta = to.getAnnotation(ExcelMeta.class);
        readToInstance(toInstnace,
                sheet,
                excelMeta.startYAxis(),
                excelMeta.endYAxis());
        return toInstnace;
    }

    public static <T extends ExcelData> void readToInstance(T toInstance, Sheet sheet, final int startYAxis, final int endYAxis) {
        Map<String, Field> targets = ExcelUtils.excelTargetFields(toInstance.getClass());
        //TODO: Reference Read 기능 추가 해야함 Writer의 Reference Read 기능 살펴 볼것.
        targets.forEach((fieldName, field) -> {
            if (field.isAnnotationPresent(CellMeta.class)) {
                String value = cellMetaData(sheet, field.getAnnotation(CellMeta.class), startYAxis);
                if (value != null) value = value.trim();
                try {
                    field.set(toInstance, value);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            } else if (field.isAnnotationPresent(RowMeta.class)) {
                RowMeta rowMeta = field.getAnnotation(RowMeta.class);
                List<ExcelData> value = null;
                switch (rowMeta.rowType()) {
//                    case RANGE, X_RANDOM -> {
                    //지원 예정
//                    }
                    case Y_RANDOM -> {
                        if (ExcelUtils.isGroupBy(rowMeta)) {
                            value = mergeRowMetaData(sheet, rowMeta, startYAxis);
                        } else {
                            value = yRandomRowMetaData(sheet, rowMeta, startYAxis, endYAxis);
                        }
                    }
                }
                try {
                    field.set(toInstance, value);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        });
    }

    public static String cellMetaData(Sheet sheet, CellMeta cellMeta, final int startYAxis) {
        List<Integer> cellNum = ExcelUtils.xAxisToCellNums(cellMeta.xAxis());
        int[] yAxies = cellMeta.yAxis();
        Row row = sheet.getRow(ExcelUtils.yAxisToRownum(ExcelUtils.sumYAxis(yAxies[0], startYAxis)));

        if (row == null) return "";

        Cell cell = row.getCell(cellNum.getFirst());
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    public static List<ExcelData> yRandomRowMetaData(Sheet sheet, RowMeta rowMeta, final int startYAxis, final int endYAxis) {
        List<ExcelData> dataList = new ArrayList<>();
        int rowStartYAxis = ExcelUtils.sumYAxis(startYAxis, rowMeta.startYAxis());
        for (int i = rowStartYAxis; i <= endYAxis; i++) {
            try {
                ExcelData row = (ExcelData) rowMeta.target().getConstructor().newInstance();
                readToInstance(row, sheet, i, i);
                dataList.add(row);
            } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                     NoSuchMethodException e) {
                e.printStackTrace();
            }
        }
        return dataList;
    }

    public static List<ExcelData> mergeRowMetaData(Sheet sheet, RowMeta rowMeta, final int startYAxis) {
        int rowMetaIndex = ExcelUtils.yAxisToRownum(ExcelUtils.sumYAxis(startYAxis, rowMeta.startYAxis()));
        return sheet.getMergedRegions()
                .stream()
                .filter(range -> {
                    return range.getFirstRow() >= rowMetaIndex &&
                            range.getFirstColumn() == ExcelUtils.xAxisToCellNums(rowMeta.xAxis()).getFirst();
                })
                .map(range -> {
                    ExcelData row = null;
                    try {
                        row = (ExcelData) rowMeta.target().getConstructor().newInstance();
                        readToInstance(row, sheet, ExcelUtils.rownumToYAxis(range.getFirstRow()), ExcelUtils.rownumToYAxis(range.getLastRow()));
                    } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                             NoSuchMethodException e) {
                        e.printStackTrace();
                    }
                    return row;
                })
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
    }
}
