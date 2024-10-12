package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.ExcelMeta;
import com.pojo.poi.core.excel.annotation.RowMeta;
import com.pojo.poi.core.excel.annotation.ValueMeta;
import com.pojo.poi.core.excel.model.ExcelData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelMaster {
    public static <T extends ExcelData> T readSheet(Class<T> to, Sheet sheet) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        T toInstnace = to.getConstructor().newInstance();
        if (!to.isAnnotationPresent(ExcelMeta.class)) return toInstnace;

        ExcelMeta excelMeta = to.getAnnotation(ExcelMeta.class);
                readToInstance(toInstnace,
                        sheet,
                        ExcelUtils.yAxisToRownum(excelMeta.startYAxis()),
                        ExcelUtils.yAxisToRownum(excelMeta.endYAxis()));
        return toInstnace;
    }

    public static <T extends ExcelData> void readToInstance(T toInstance, Sheet sheet, final int startRow, final int endRow) {
        Map<String, Field> targets = excelTargetFields(toInstance.getClass());
        targets.forEach((fieldName, field) -> {
            if (field.isAnnotationPresent(CellMeta.class)) {
                String value = cellMetaData(sheet, field.getAnnotation(CellMeta.class), startRow);
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
                    case RANGE -> {
                        //지원 예정
                    }
                    case Y_RANDOM -> {
                        if (isGroupBy(rowMeta)) {
                            value = yRandomRowMetaData(sheet, rowMeta, startRow);
                        } else {
                            value = mergeRowMetaData(sheet, rowMeta);
                        }
                    }
                    case X_RANDOM -> {
                        //지원 예정
                    }
                }
            }
        });
    }

    public static String cellMetaData(Sheet sheet, CellMeta cellMeta, final int startRow) {
        ValueMeta headerMeta = cellMeta.headerMeta();
        List<Integer> cellNum = ExcelUtils.xAxisToCellNums(cellMeta.xAxis());
        List<Integer> rownum = ExcelUtils.yAxisToRownums(cellMeta.yAxis());
        Row row = sheet.getRow(rownum.getFirst() + startRow);

        if (row == null) return "";

        Cell cell = row.getCell(cellNum.getFirst());
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    public static List<ExcelData> yRandomRowMetaData(Sheet sheet, RowMeta rowMeta, final int startRow) {
        List<ExcelData> dataList = new ArrayList<>();
        int rowMetaIndex = startRow + ExcelUtils.yAxisToRownum(rowMeta.startYAxis());
        int endRowIndex = sheet.getLastRowNum();
        for (int i = rowMetaIndex; i < endRowIndex; i++) {
            try {
                ExcelData row = (ExcelData) rowMeta.target().getConstructor().newInstance();
                readToInstance(row, sheet, i, endRowIndex);
                dataList.add(row);
            } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                     NoSuchMethodException e) {
//                throw new RuntimeException(e);
                e.printStackTrace();
            }
        }
        return dataList;
    }

    public static List<ExcelData> mergeRowMetaData(Sheet sheet, RowMeta rowMeta) {
        return sheet.getMergedRegions()
                .stream()
                .filter(range -> {
                    return range.getFirstRow() >= ExcelUtils.yAxisToRownum(rowMeta.startYAxis()) &&
                            range.getFirstColumn() == ExcelUtils.xAxisToCellNums(rowMeta.xAxis()).getFirst();
                })
                .map(range -> {
                    ExcelData row = null;
                    try {
                        row = (ExcelData) rowMeta.target().getConstructor().newInstance();
                        readToInstance(row, sheet, range.getFirstRow(), range.getLastRow());
                    } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                             NoSuchMethodException e) {
//                        throw new RuntimeException(e);
                        e.printStackTrace();
                    }
                    return row;
                })
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
    }


    public static <T> Map<String, Field> excelTargetFields(Class<T> to) {
        Map<String, Field> targets = new HashMap<>();
        for (Field field : to.getDeclaredFields()) {
            field.setAccessible(true);
            if (field.isAnnotationPresent(CellMeta.class) || field.isAnnotationPresent(RowMeta.class)) {
                targets.put(field.getName(), field);
            }
        }
        return targets;
    }

    public static boolean isGroupBy(RowMeta rowMeta) {
        return rowMeta.groupBys().length > 0;
    }
}
