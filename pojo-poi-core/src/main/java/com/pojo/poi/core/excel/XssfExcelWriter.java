package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.*;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelRowStyle;
import com.pojo.poi.core.excel.style.ExcelStyleManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.*;

public class XssfExcelWriter {
    public static void writeExcelData(final Sheet sheet, final ExcelStyleManager excelStyleManager, ExcelData excelData, final int startYAxis) {
        if (!excelData.getClass().isAnnotationPresent(ExcelMeta.class)) return;

        Map<String, Field> targetFields = ExcelUtils.excelTargetFields(excelData.getClass());

        ExcelMeta excelMeta = excelData.getClass().getAnnotation(ExcelMeta.class);
        final int excelMetaStartYAxis = ExcelUtils.sumYAxis(startYAxis, excelMeta.startYAxis());
        ValueMeta[] headerMetas = excelMeta.headerMetas();
        for (ValueMeta headerMeta : headerMetas) {
            writeValueMeta(sheet, excelStyleManager, headerMeta, excelMetaStartYAxis);
        }
        Map<Integer, ExcelDataMeta> excelMetaDataMap = new HashMap<>();
        targetFields.values().stream()
                .filter(field -> field.isAnnotationPresent(CellMeta.class) || field.isAnnotationPresent(RowMeta.class))
                .sorted((f1, f2) -> {
                    MetaOrder metaOrderF1 = ExcelUtils.metaOrder(f1);
                    MetaOrder metaOrderF2 = ExcelUtils.metaOrder(f2);

                    return Integer.compare(Objects.requireNonNull(metaOrderF1).value(), Objects.requireNonNull(metaOrderF2).value());
                })
                .map(field -> {
                    ExcelDataMeta excelDataMeta = new ExcelDataMeta(field);
                    excelMetaDataMap.put(excelDataMeta.getMetaOrder().value(), excelDataMeta);
                    return excelDataMeta;
                })
                .forEach(excelDataMeta -> {
                    //TODO: XAxis 에 대한 값도 추가해 줘야함 X_RANDOM 기능 만들때 고려하자.
                    int yAxis = findReferenceYAxis(excelMetaDataMap, excelMetaStartYAxis, excelDataMeta);
                    if (excelDataMeta.isCellMeta()) {
                        writeCellMeta(excelDataMeta, sheet, excelStyleManager, excelData, yAxis);
                    }
                    else if (excelDataMeta.isRowMeta()) {
                        writeRowMeta(excelDataMeta, sheet, excelStyleManager, excelData, yAxis);
                    }
                });
    }

    private static int findReferenceYAxis(Map<Integer, ExcelDataMeta> excelMetaDataMap, int currentYAxis,  ExcelDataMeta excelDataMeta) {
        switch (excelDataMeta.getMetaOrder().type()) {
            case NONE -> {
                return currentYAxis;
            }
            case Y_REFERENCES -> {
                ExcelDataMeta reference = excelMetaDataMap.get(excelDataMeta.getMetaOrder().referenceMetaOrder());
                //마지막 라인 + 1
                return reference != null ? reference.getEndYAxis() + 1 : currentYAxis;
            }
        }
        return currentYAxis;
    }

    private static void writeCellMeta(final ExcelDataMeta excelDataMeta, final Sheet sheet, final ExcelStyleManager excelStyleManager, ExcelData excelData, final int startYAxis) {
        CellMeta cellMeta = excelDataMeta.getField().getAnnotation(CellMeta.class);
        Object data = null;
        try {
            data = excelDataMeta.getField().get(excelData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        writeCellData(sheet, excelStyleManager, cellMeta, startYAxis, data);
    }

    private static void writeRowMeta(final ExcelDataMeta excelDataMeta, final Sheet sheet, final ExcelStyleManager excelStyleManager, ExcelData excelData, final int startYAxis) {
        RowMeta rowMeta = excelDataMeta.getField().getAnnotation(RowMeta.class);
        int rowMetaStartYAxis = ExcelUtils.sumYAxis(startYAxis, rowMeta.startYAxis());
        excelDataMeta.setStartYAxis(rowMetaStartYAxis);
        ValueMeta[] rowHeaderMetas = rowMeta.headerMetas();
        int maxValueMetaYAxis = 0;
        for (ValueMeta rowHeaderMeta : rowHeaderMetas) {
            writeValueMeta(sheet, excelStyleManager, rowHeaderMeta, rowMetaStartYAxis);
            int[] rowHeaderMetaYAxis = rowHeaderMeta.yAxis();
            Arrays.sort(rowHeaderMetaYAxis);
            maxValueMetaYAxis = Math.max(maxValueMetaYAxis, rowHeaderMetaYAxis[rowHeaderMetaYAxis.length - 1]);
        }
        rowMetaStartYAxis = ExcelUtils.sumYAxis(rowMetaStartYAxis, maxValueMetaYAxis + 1);
        try {
            if (!Collection.class.isAssignableFrom(excelDataMeta.getField().getType())) {
                throw new RuntimeException(String.format("필드 %s 는 컬렉션이 아닙니다., RowMeta 는 컬렉션에 적용 가능합니다.", excelDataMeta.getField().getName()));
            }
            Collection<ExcelData> innerExcelDatas = (Collection<ExcelData>) excelDataMeta.getField().get(excelData);
            if (innerExcelDatas == null) {
                System.out.printf("filed data is null, filed name: %s", excelDataMeta.getField().getName());
                return;
            }
            //Row Meta 데이터를 먼저 쓰고 난 후 머지를 한다.
            Iterator<ExcelData> iterator = innerExcelDatas.iterator();
            int lastYAxis = ExcelUtils.rownumToYAxis(sheet.getLastRowNum());
            for (
                    int i = 0, firstYAxis = rowMetaStartYAxis;
                    iterator.hasNext();
                    i++, firstYAxis = lastYAxis + 1
            ) {
                writeExcelData(sheet, excelStyleManager, iterator.next(), firstYAxis);
                //merge 할 마지막 row
                lastYAxis = ExcelUtils.rownumToYAxis(sheet.getLastRowNum());
                if (ExcelUtils.isGroupBy(rowMeta)) {
                    if (firstYAxis == lastYAxis) continue;
                    GroupByMeta[] groupByMetas = rowMeta.groupBys();
                    for (GroupByMeta groupByMeta : groupByMetas) {
                        int[] yAxes = {
                                ExcelUtils.sumYAxis(firstYAxis, groupByMeta.yAxis()[0]),
                                lastYAxis
                        };
                        switch (groupByMeta.dataType()) {
                            case AUTO_INCREMENT -> {
                                writeGroupBy(sheet, excelStyleManager, groupByMeta, groupByMeta.xAxis(), yAxes, i + 1);
                            }
                            case CELL_DATA -> {
                                writeGroupBy(sheet, excelStyleManager, groupByMeta, groupByMeta.xAxis(), yAxes);
                            }
                        }
                    }
                }
            }
            //Row Style
            applyRowStyle(sheet, excelStyleManager, rowMetaStartYAxis, lastYAxis, rowMeta.rowStyle());
            excelDataMeta.setEndYAxis(lastYAxis);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    //TODO: Gruop By 내에서 write 를 하게 되면 중복이 발생함
    //  현재 로직 기준으로 cellMeta, RowMeta 데이터를 모두 Excel 에 기록 후 Group By 를 하게 되기 때문
    public static void writeGroupBy(final Sheet sheet, final ExcelStyleManager excelStyleManager, GroupByMeta groupByMeta, String[] xAxes, int[] yAxes) {
        //TODO: Group By Type 별 분기 추가 하기.
        if (xAxes == null) xAxes = groupByMeta.xAxis();
        List<String> fromToXAxesList = ExcelUtils.xAxisFromToxAxis(xAxes);
        String[] fromToXAxes = new String[fromToXAxesList.size()];
        fromToXAxesList.toArray(fromToXAxes);
        List<Integer> fromToYAxesList = ExcelUtils.yAxisFromToyAxis(yAxes);
        int[] fromToYAxes = new int[fromToYAxesList.size()];
        for (int i = 0; i < fromToYAxesList.size(); i++) {
            fromToYAxes[i] = fromToYAxesList.get(i);
        }
        prepareRegion(sheet, excelStyleManager, fromToXAxes, fromToYAxes, groupByMeta.cellStyle());
        Arrays.sort(xAxes);
        Arrays.sort(yAxes);
        cellMerging(sheet, fromToXAxes, fromToYAxes);
        writeToCell(sheet, xAxes[0], yAxes[0], ExcelUtils.cell(ExcelUtils.row(sheet, yAxes[0]), xAxes[0]).getStringCellValue());
    }

    //TODO: Gruop By 내에서 write 를 하게 되면 중복이 발생함
    //  현재 로직 기준으로 cellMeta, RowMeta 데이터를 모두 Excel 에 기록 후 Group By 를 하게 되기 때문
    public static void writeGroupBy(final Sheet sheet, final ExcelStyleManager excelStyleManager, GroupByMeta groupByMeta, String[] xAxes, int[] yAxes, Object data) {
        //TODO: Group By Type 별 분기 추가 하기.
        if (xAxes == null) xAxes = groupByMeta.xAxis();
        prepareRegion(sheet, excelStyleManager, xAxes, yAxes, groupByMeta.cellStyle());
        Arrays.sort(xAxes);
        Arrays.sort(yAxes);
        cellMerging(sheet, xAxes, yAxes);
        writeToCell(sheet, xAxes[0], yAxes[0], ExcelUtils.toStringData(data));
    }

    public static void writeCellData(final Sheet sheet, final ExcelStyleManager excelStyleManager, CellMeta cellMeta, final int startYAxis, Object data) {
        writeValueMeta(sheet, excelStyleManager, cellMeta.headerMeta(), startYAxis);
        String[] xAxes = cellMeta.xAxis();
        int[] yAxes = new int[cellMeta.yAxis().length];
        for (int i = 0; i < yAxes.length; i++) {
            yAxes[i] = ExcelUtils.sumYAxis(cellMeta.yAxis()[i], startYAxis);
        }
        prepareRegion(sheet, excelStyleManager, xAxes, yAxes, cellMeta.cellStyle());
        Arrays.sort(xAxes);
        Arrays.sort(yAxes);
        if (xAxes.length > 1 || yAxes.length > 1) {
            cellMerging(sheet, xAxes, yAxes);
        }
        writeToCell(sheet, xAxes[0], yAxes[0], ExcelUtils.toStringData(data));
    }

    public static void writeValueMeta(final Sheet sheet, final ExcelStyleManager excelStyleManager, ValueMeta valueMeta, final int startYAxis) {
        if (valueMeta.value().isEmpty()) return;

        String[] xAxes = valueMeta.xAxis();
        int[] yAxes = new int[valueMeta.yAxis().length];
        for (int i = 0; i < yAxes.length; i++) {
            yAxes[i] = ExcelUtils.sumYAxis(valueMeta.yAxis()[i], startYAxis);
        }

        prepareRegion(sheet, excelStyleManager, xAxes, yAxes, valueMeta.cellStyle());
        Arrays.sort(xAxes);
        Arrays.sort(yAxes);
        if (xAxes.length > 1 || yAxes.length > 1) {
            cellMerging(sheet, xAxes, yAxes);
        }
        writeToCell(sheet, xAxes[0], yAxes[0], valueMeta.value());
    }

    public static void cellMerging(final Sheet sheet, String[] xAxes, int[] yAxes) {
        Arrays.sort(xAxes);
        Arrays.sort(xAxes);
        if (!ExcelUtils.isMergedCell(xAxes, yAxes)) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(
                ExcelUtils.yAxisToRownum(yAxes[0]),
                ExcelUtils.yAxisToRownum(yAxes[yAxes.length - 1]),
                ExcelUtils.xAxisToCellNum(xAxes[0]),
                ExcelUtils.xAxisToCellNum(xAxes[xAxes.length - 1])
        ));
    }

    //TODO: 데이터는 아직 String 만 지원, 여러 데이터 형을 지원할 필요가 있을까?
    public static void writeToCell(final Sheet sheet, String xAxis, int yAxis, String data) {
//        System.out.printf("%s:%s", xAxis + yAxis, data);
//        System.out.println();
        Cell cell = ExcelUtils.cell(ExcelUtils.row(sheet, yAxis), xAxis);
        cell.setCellValue(data);
    }

    public static void prepareRegion(final Sheet sheet, final ExcelStyleManager excelStyleManager, String[] xAxes, int[] yAxes, ExcelCellStyle[] cellStyle) {
        for (int yAxis : yAxes) {
            Row row = ExcelUtils.row(sheet, yAxis);
            for (String xAxis : xAxes) {
                Cell cell = ExcelUtils.cell(row, xAxis);
                applyCellStyle(excelStyleManager, cell, cellStyle);
            }
        }
    }

    public static void applyCellStyle(final ExcelStyleManager excelStyleManager, Cell cell, ExcelCellStyle[] excelCellStyles) {
        if (excelCellStyles.length < 1) return;
        excelStyleManager.applyCellStyle(cell, excelCellStyles);
    }

    //TODO: row style 은 고민해 보기
    public static void applyRowStyle(final Sheet sheet, final ExcelStyleManager excelStyleManager, int startYAxis, int endYAxis, ExcelRowStyle[] excelRowStyles) {
        if (excelRowStyles.length < 1) return;
        int[] yAxes = {
                startYAxis,
                endYAxis
        };
        List<Integer> fromToYAxesList = ExcelUtils.yAxisFromToyAxis(yAxes);
        List<Row> fromToRows = fromToYAxesList.stream()
                .map(yAxis -> ExcelUtils.row(sheet, yAxis))
                .toList();
        excelStyleManager.applyRowStyle(fromToRows, excelRowStyles);
    }
}
