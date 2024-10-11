package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.*;
import com.pojo.poi.core.excel.model.ExcelCellStyle;
import com.pojo.poi.core.excel.model.ExcelData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelModel {
    private static final String DEFAULT_SHEET_NAME = "sheet";
    private final Map<ExcelCellStyle, CellStyle> cellStyles = new HashMap<>();
    private final Map<ExcelCellStyle, Font> fonts = new HashMap<>();
    private final XSSFWorkbook workbook;
    private final String originFileName;
    private final List<ExcelSheetModel> sheets = new ArrayList<>();
    private String fileName;
    private int sheetIndex = 1;

    public ExcelModel(XSSFWorkbook workbook, String originFileName, String fileName, int sheetIndex) {
        this.workbook = workbook;
        this.originFileName = originFileName;
        this.fileName = fileName;
        this.sheetIndex = sheetIndex;
    }

    public static ExcelModelBuilder builder(String fileName) {
        if (!fileName.endsWith(".xlsx")) fileName = fileName.concat(".xlsx");
        return ExcelModelBuilder.builder(fileName);
    }

    public ExcelModel fileName(String fileName) {
        if (!fileName.endsWith(".xlsx")) fileName = fileName.concat(".xlsx");
        this.fileName = fileName;
        return this;
    }

    public ExcelModel addExcelDatas(List<ExcelData> excelDatas) {
        String sheetName = DEFAULT_SHEET_NAME + sheetIndex++;
        return addExcelDatas(sheetName, excelDatas, null);
    }

    public ExcelModel addExcelDatas(final String sheetName, List<ExcelData> excelDatas, float[] cellWidths) {
        Optional<ExcelSheetModel> sheetModel = this.sheets.stream()
                .filter(excelSheetModel -> excelSheetModel.sheetName.equals(sheetName))
                .findAny();
        sheetModel.ifPresentOrElse(excelSheetModel -> excelSheetModel.excelDatas.addAll(excelDatas), () -> {
            ExcelSheetModel excelSheetModel = new ExcelSheetModel(this.workbook, sheetName, cellWidths, this.cellStyles, this.fonts);
            excelSheetModel.excelDatas.addAll(excelDatas);
            this.sheets.add(excelSheetModel);
        });
        return this;
    }

    public ExcelModel writeAll() {
        this.sheets.forEach(ExcelSheetModel::write);
        return this;
    }

    public ExcelModel end() {
        return this;
    }

    public ByteArrayInputStream getExcelStream() {
        try (ByteArrayOutputStream bao = new ByteArrayOutputStream()) {
            this.workbook.write(bao);
            return new ByteArrayInputStream(bao.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        } finally {
            // workbook 이 SXSSFWorkBook 일때, 디스크에 임시로 저장한 파일 삭제함
            // this.workbook.dispose();
            try {
                this.workbook.close();
            } catch (IOException ignored) {
            }
        }
    }

    private static class ExcelSheetModel {
        Map<ExcelCellStyle, CellStyle> cellStyleMap;
        Map<ExcelCellStyle, Font> fontMap;
        private final float[] defaultCellWiths = {};
        private final Sheet sheet;
        private final String sheetName;
        private final List<ExcelData> excelDatas;

        public ExcelSheetModel(Workbook workbook,
                               String sheetName,
                               float[] cellWidths,
                               Map<ExcelCellStyle, CellStyle> cellStyleMap,
                               Map<ExcelCellStyle, Font> fontMap) {
            this.sheet = this.cresheet(workbook, sheetName);
            this.sheetName = sheetName;
            this.excelDatas = new ArrayList<>();
            this.cellStyleMap = cellStyleMap;
            this.fontMap = fontMap;
            this.setColumnWidths(cellWidths);

        }

        private Sheet cresheet(Workbook workbook, String sheetName) {
            return workbook.createSheet(sheetName);
        }

        /**
         * Excel Cell Width 값을 Poi Width 값으로 변환
         *
         * @param excelCellWith excel 상 width 값
         * @return Poi Width 값
         */
        private int width256(float excelCellWith) {
            return (int) Math.floor((excelCellWith * Units.DEFAULT_CHARACTER_WIDTH + 5) / Units.DEFAULT_CHARACTER_WIDTH * 256);
        }

        public void setColumnWidths(float[] cellWidths) {
            if (cellWidths == null || cellWidths.length == 0) cellWidths = this.defaultCellWiths;
            for (int i = 0; i < cellWidths.length; i++) {
                sheet.setColumnWidth(i, width256(cellWidths[i]));
            }
        }

        /**
         * Excel 에 데이터를 기록
         */
        public void write() {
            this.write(this.excelDatas);
        }

        public void write(List<? extends ExcelData> excelDatas) {
            this.write(excelDatas, 0);
        }

        public void write(List<? extends ExcelData> excelDatas, int startRow) {
            excelDatas.forEach(excelData -> this.writeExcelData(excelData, startRow));
        }

        public void writeExcelData(ExcelData excelData, final int startRow) {
            if (!excelData.getClass().isAnnotationPresent(ModelMeta.class)) return;

            Map<String, Field> targetFields = ExcelMaster.excelTargetFields(excelData.getClass());

            ModelMeta modelMeta = excelData.getClass().getAnnotation(ModelMeta.class);
            int modelMetaStartRow = ExcelUtils.yAxisToRownum(modelMeta.startYAxis()) + startRow;
            ValueMeta[] headerMetas = modelMeta.headerMetas();
            for (ValueMeta headerMeta : headerMetas) {
                writeValueMeta(this.sheet, headerMeta, modelMetaStartRow);
            }
            //TODO: CellMeta, RowMeta 기준이 아닌 Y_RANDOM 기준으로 바꿔야 함.
            //TODO: Y_RANDOM 의 경우 시작 row 값을 알 수 없으므로 불가능함.
            targetFields.values().stream()
                    .filter(field -> field.isAnnotationPresent(CellMeta.class))
                    .forEach(field -> {
                        CellMeta cellMeta = field.getAnnotation(CellMeta.class);
                        Object data = null;
                        try {
                            data = field.get(excelData);
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        }
                        writeCellData(this.sheet, cellMeta, startRow, data);
                    });
            targetFields.values().stream()
                    .filter(field -> field.isAnnotationPresent(RowMeta.class))
                    .forEach(field -> {
                        int fieldStartRow = startRow;
                        RowMeta rowMeta = field.getAnnotation(RowMeta.class);
                        ValueMeta[] rowHeaderMetas = rowMeta.headerMetas();
                        for (ValueMeta rowHeaderMeta : rowHeaderMetas) {
                            writeValueMeta(this.sheet, rowHeaderMeta, fieldStartRow);
                        }
                        fieldStartRow += ExcelUtils.yAxisToRownum(rowMeta.startYAxis());
                        try {
                            List<ExcelData> innerExcelDatas = (List<ExcelData>) field.get(excelData);
                            if (innerExcelDatas == null) {
                                System.out.printf("filed data is null, filed name: %s", field.getName());
                                return;
                            }
                            for (int i = 0, listStartRow = fieldStartRow, listLastRow; i < innerExcelDatas.size(); i++, listStartRow = listLastRow + 1) {
                                writeExcelData(innerExcelDatas.get(i), listStartRow);
                                listLastRow = this.sheet.getLastRowNum();
                                if (rowMeta.isGroupBy()) {
                                    if (listStartRow == listLastRow) continue;
                                    GroupByMeta[] groupByMetas = rowMeta.groupBys();
                                    for (GroupByMeta groupByMeta : groupByMetas) {
                                        int[] yAxes = {
                                                ExcelUtils.rownumToYAxis(listLastRow) + groupByMeta.yAxis()[0] - 1,
                                                ExcelUtils.rownumToYAxis(listLastRow)
                                        };
                                        switch (groupByMeta.dataType()) {
                                            case AUTO_INCREMENT -> {
                                                writeGroupBy(this.sheet, groupByMeta, groupByMeta.xAxis(), yAxes, i + 1);
                                            }
                                            case CELL_DATA -> {
                                                writeGroupBy(this.sheet, groupByMeta, groupByMeta.xAxis(), yAxes);
                                            }
                                        }
                                    }
                                }
                            }
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        }
                    });
        }

        public void writeGroupBy(Sheet sheet, GroupByMeta groupByMeta, String[] xAxes, int[] yAxes) {
            //TODO: Group By Type 별 분기 추가 하기.
            if (xAxes == null) xAxes = groupByMeta.xAxis();
            List<String> fromToXAxesList = ExcelUtils.xAxisFromToxAxis(xAxes);
            String[] fromToXAxes = new String[fromToXAxesList.size()];
            fromToXAxesList.toArray(fromToXAxes);
            List<Integer> fromToYAxesList = ExcelUtils.yAxisFromToxAxis(yAxes);
            int[] fromToYAxes = new int[fromToYAxesList.size()];
            for (int i = 0; i < fromToYAxesList.size(); i++) {
                fromToYAxes[i] = fromToYAxesList.get(i);
            }
            prepareRegion(sheet, fromToXAxes, fromToYAxes, groupByMeta.cellStyle());
            mergingCells(sheet, fromToXAxes, fromToYAxes, cell(row(sheet, yAxes[0]), xAxes[0]).getStringCellValue());
        }

        public void writeGroupBy(Sheet sheet, GroupByMeta groupByMeta, String[] xAxes, int[] yAxes, Object data) {
            //TODO: Group By Type 별 분기 추가 하기.
            if (xAxes == null) xAxes = groupByMeta.xAxis();
            prepareRegion(sheet, xAxes, yAxes, groupByMeta.cellStyle());
            mergingCells(sheet, xAxes, yAxes, data);
            writeToCell(sheet, xAxes[0], yAxes[0], toStringData(data));
        }

        public void writeCellData(Sheet sheet, CellMeta cellMeta, final int startRow, Object data) {
            writeValueMeta(sheet, cellMeta.headerMeta(), startRow);
            String[] xAxes = cellMeta.xAxis();
            int[] yAxes = new int[cellMeta.yAxis().length];
            switch (cellMeta.axisType()) {
                case NORMAL -> {
                }
                case Y_RANDOM -> {
                    for (int i = 0; i < yAxes.length; i++) {
                        yAxes[i] = ExcelUtils.yAxisToRownum(cellMeta.yAxis()[i]) + ExcelUtils.rownumToYAxis(startRow);
                    }
                }
                case X_RANDOM -> {
                }
            }
            prepareRegion(sheet, xAxes, yAxes, cellMeta.cellStyle());
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            if (xAxes.length > 1 || yAxes.length > 1) {
                mergingCells(sheet, xAxes, yAxes, toStringData(data));
            }
        }

        public void writeValueMeta(Sheet sheet, ValueMeta valueMeta, final int startRow) {
            if (valueMeta.value().isEmpty()) return;

            String[] xAxes = valueMeta.xAxis();
            int[] yAxes = new int[valueMeta.yAxis().length];
            for (int i = 0; i < yAxes.length; i++) {
                yAxes[i] = ExcelUtils.yAxisToRownum(valueMeta.yAxis()[i]) + ExcelUtils.rownumToYAxis(startRow);
            }

            prepareRegion(sheet, xAxes, yAxes, valueMeta.cellStyle());
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            if (xAxes.length > 1 || yAxes.length > 1) {
                mergingCells(sheet, xAxes, yAxes, valueMeta.value());
            }
            writeToCell(sheet, xAxes[0], yAxes[0], valueMeta.value());
        }

        //TODO: 데이터는 아직 String 만 지원, 여러 데이터 형을 지원할 필요가 있을까?
        public void mergingCells(Sheet sheet, String[] xAxes, int[] yAxes, Object data) {
            for (String xAxis : xAxes) {
                for (int yAxis : yAxes) {
                    if (data != null) {
                        writeToCell(sheet, xAxis, yAxis, toStringData(data));
                    }
                }
            }
            Arrays.sort(xAxes);
            Arrays.sort(xAxes);
            if (xAxes.length == 1 && yAxes.length == 1) return;
            sheet.addMergedRegion(new CellRangeAddress(
                    ExcelUtils.yAxisToRownum(yAxes[0]),
                    ExcelUtils.yAxisToRownum(yAxes[yAxes.length - 1]),
                    ExcelUtils.xAxisToCellNum(xAxes[0]),
                    ExcelUtils.xAxisToCellNum(xAxes[xAxes.length - 1])
            ));
        }

        //TODO: 데이터는 아직 String 만 지원, 여러 데이터 형을 지원할 필요가 있을까?
        //TODO: 데이터를 중복으로 쓰는 문제가 있음. 로그 확인하여 바꿀 필요가 있음
        public void writeToCell(Sheet sheet, String xAxis, int yAxis, String data) {
            System.out.printf("write to cell...X열: %s, Y열: %s, data: %s%n", xAxis, yAxis, data);
            cell(row(sheet, yAxis), xAxis).setCellValue(data);
        }

        public void prepareRegion(Sheet sheet, String[] xAxes, int[] yAxes, ExcelCellStyle cellStyle) {
            for (int yAxis : yAxes) {
                Row row = row(sheet, yAxis);
                for (String xAxis : xAxes) {
                    Cell cell = cell(row, xAxis);
                    if (cellStyle != null) {
                        cellStyle.applyCellStyle(cell, this.cellStyleMap, this.fontMap);
                    }
                }
            }
        }

        public Row row(Sheet sheet, int yAxis) {
            int rowNum = ExcelUtils.yAxisToRownum(yAxis);
            Row row = sheet.createRow(rowNum);
            if (row == null) row = sheet.createRow(rowNum);
            return row;
        }

        public Cell cell(Row row, String xAxis) {
            int cellNum = ExcelUtils.xAxisToCellNum(xAxis);
            Cell cell = row.getCell(cellNum);
            if (cell == null) cell = row.createCell(cellNum);
            return cell;
        }

        public String toStringData(Object data) {
            return data == null ? "" : data.toString();
        }
    }

    public static class ExcelModelBuilder {
        private XSSFWorkbook workbook;
        private String originFileName;
        private String fileName;

        private static ExcelModelBuilder builder(String excelFileName) {
            ExcelModelBuilder builder = new ExcelModelBuilder();
            builder.originFileName = excelFileName;
            builder.workbook = new XSSFWorkbook();
            builder.fileName = excelFileName;
            return builder;
        }

        public ExcelModelBuilder setWorkbook(XSSFWorkbook workbook) {
            this.workbook = workbook;
            return this;
        }

        public ExcelModelBuilder setOriginFileName(String originFileName) {
            this.originFileName = originFileName;
            return this;
        }

        public ExcelModelBuilder setFileName(String fileName) {
            this.fileName = fileName;
            return this;
        }

        public ExcelModel build() {
            return new ExcelModel(this.workbook, this.originFileName, this.fileName, 1);
        }
    }
}
