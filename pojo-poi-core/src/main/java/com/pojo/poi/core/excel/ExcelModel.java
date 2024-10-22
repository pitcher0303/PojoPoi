package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.*;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

@Getter
public class ExcelModel {
    private static final String DEFAULT_SHEET_NAME = "sheet";
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
            ExcelSheetModel excelSheetModel = new ExcelSheetModel(this.workbook, sheetName, cellWidths);
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
        private final float[] defaultCellWiths = {};
        private final Sheet sheet;
        private final String sheetName;
        private final List<ExcelData> excelDatas;

        public ExcelSheetModel(Workbook workbook,
                               String sheetName,
                               float[] cellWidths) {
            this.sheet = this.cresheet(workbook, sheetName);
            this.sheetName = sheetName;
            this.excelDatas = new ArrayList<>();
            this.setColumnWidths(cellWidths);

        }

        private Sheet cresheet(Workbook workbook, String sheetName) {
            return workbook.createSheet(sheetName);
        }

        public void setColumnWidths(float[] cellWidths) {
            if (cellWidths == null || cellWidths.length == 0) cellWidths = this.defaultCellWiths;
            for (int i = 0; i < cellWidths.length; i++) {
                sheet.setColumnWidth(i, ExcelUtils.width256(cellWidths[i]));
            }
        }

        /**
         * Excel 에 데이터를 기록
         */
        public void write() {
            this.write(this.excelDatas);
        }

        public void write(List<? extends ExcelData> excelDatas) {
            this.write(excelDatas, 1);
        }

        public void write(List<? extends ExcelData> excelDatas, int startYAxis) {
            excelDatas.forEach(excelData -> this.writeExcelData(excelData, startYAxis));
        }

        public void writeExcelData(ExcelData excelData, final int startYAxis) {
            if (!excelData.getClass().isAnnotationPresent(ExcelMeta.class)) return;

            Map<String, Field> targetFields = ExcelMaster.excelTargetFields(excelData.getClass());

            ExcelMeta excelMeta = excelData.getClass().getAnnotation(ExcelMeta.class);
            ValueMeta[] headerMetas = excelMeta.headerMetas();
            for (ValueMeta headerMeta : headerMetas) {
                writeValueMeta(this.sheet, headerMeta, ExcelUtils.sumYAxis(excelMeta.startYAxis(), startYAxis));
            }
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
                        writeCellData(this.sheet, cellMeta, startYAxis, data);
                    });
            targetFields.values().stream()
                    .filter(field -> field.isAnnotationPresent(RowMeta.class))
                    .forEach(field -> {
                        RowMeta rowMeta = field.getAnnotation(RowMeta.class);
                        ValueMeta[] rowHeaderMetas = rowMeta.headerMetas();
                        for (ValueMeta rowHeaderMeta : rowHeaderMetas) {
                            writeValueMeta(this.sheet, rowHeaderMeta, startYAxis);
                        }
                        try {
                            if(!Collection.class.isAssignableFrom(field.getType())) {
                                throw new RuntimeException(String.format("필드 %s 는 컬렉션이 아닙니다., RowMeta 는 컬렉션에 적용 가능합니다.", field.getName()));
                            }
                            Collection<ExcelData> innerExcelDatas = (Collection<ExcelData>) field.get(excelData);
                            if (innerExcelDatas == null) {
                                System.out.printf("filed data is null, filed name: %s", field.getName());
                                return;
                            }
                            //Row Meta 데이터를 먼저 쓰고 난 후 머지를 한다.
                            Iterator<ExcelData> iterator = innerExcelDatas.iterator();
                            for (
                                    int i = 0, firstYAxis = ExcelUtils.sumYAxis(startYAxis, rowMeta.startYAxis()), lastYAxis;
                                    iterator.hasNext();
                                    i++, firstYAxis = lastYAxis + 1
                            ) {
                                writeExcelData(iterator.next(), firstYAxis);
                                //merge 할 마지막 row
                                lastYAxis = ExcelUtils.rownumToYAxis(this.sheet.getLastRowNum());
                                if (ExcelMaster.isGroupBy(rowMeta)) {
                                    if (firstYAxis == lastYAxis) continue;
                                    GroupByMeta[] groupByMetas = rowMeta.groupBys();
                                    for (GroupByMeta groupByMeta : groupByMetas) {
                                        int[] yAxes = {
                                                ExcelUtils.sumYAxis(firstYAxis, groupByMeta.yAxis()[0]),
                                                lastYAxis
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

        //TODO: Gruop By 내에서 write 를 하게 되면 중복이 발생함
        //  현재 로직 기준으로 cellMeta, RowMeta 데이터를 모두 Excel 에 기록 후 Group By 를 하게 되기 때문
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
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            cellMerging(sheet, fromToXAxes, fromToYAxes);
            writeToCell(sheet, xAxes[0], yAxes[0], cell(row(sheet, yAxes[0]), xAxes[0]).getStringCellValue());

        }

        //TODO: Gruop By 내에서 write 를 하게 되면 중복이 발생함
        //  현재 로직 기준으로 cellMeta, RowMeta 데이터를 모두 Excel 에 기록 후 Group By 를 하게 되기 때문
        public void writeGroupBy(Sheet sheet, GroupByMeta groupByMeta, String[] xAxes, int[] yAxes, Object data) {
            //TODO: Group By Type 별 분기 추가 하기.
            if (xAxes == null) xAxes = groupByMeta.xAxis();
            prepareRegion(sheet, xAxes, yAxes, groupByMeta.cellStyle());
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            cellMerging(sheet, xAxes, yAxes);
            writeToCell(sheet, xAxes[0], yAxes[0], toStringData(data));
        }

        public void writeCellData(Sheet sheet, CellMeta cellMeta, final int startYAxis, Object data) {
            writeValueMeta(sheet, cellMeta.headerMeta(), startYAxis);
            String[] xAxes = cellMeta.xAxis();
            int[] yAxes = new int[cellMeta.yAxis().length];
            for (int i = 0; i < yAxes.length; i++) {
                yAxes[i] = ExcelUtils.sumYAxis(cellMeta.yAxis()[i], startYAxis);
            }
            prepareRegion(sheet, xAxes, yAxes, cellMeta.cellStyle());
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            if (xAxes.length > 1 || yAxes.length > 1) {
                cellMerging(sheet, xAxes, yAxes);
            }
            writeToCell(sheet, xAxes[0], yAxes[0], toStringData(data));
        }

        public void writeValueMeta(Sheet sheet, ValueMeta valueMeta, final int startYAxis) {
            if (valueMeta.value().isEmpty()) return;

            String[] xAxes = valueMeta.xAxis();
            int[] yAxes = new int[valueMeta.yAxis().length];
            for (int i = 0; i < yAxes.length; i++) {
                yAxes[i] = ExcelUtils.sumYAxis(valueMeta.yAxis()[i], startYAxis);
            }

            prepareRegion(sheet, xAxes, yAxes, valueMeta.cellStyle());
            Arrays.sort(xAxes);
            Arrays.sort(yAxes);
            if (xAxes.length > 1 || yAxes.length > 1) {
                cellMerging(sheet, xAxes, yAxes);
            }
            writeToCell(sheet, xAxes[0], yAxes[0], valueMeta.value());
        }

        public void cellMerging(Sheet sheet, String[] xAxes, int[] yAxes) {
            Arrays.sort(xAxes);
            Arrays.sort(xAxes);
            if (xAxes.length == 1 && yAxes.length == 1) {
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
        public void writeToCell(Sheet sheet, String xAxis, int yAxis, String data) {
            Cell cell = cell(row(sheet, yAxis), xAxis);
            cell.setCellValue(data);
            CellStyle cellStyle = cell.getCellStyle();
            if(cellStyle == null) {
                cellStyle = sheet.getWorkbook().createCellStyle();
                cell.setCellStyle(cellStyle);
            }
            cellStyle.setWrapText(true);
        }

        public void prepareRegion(Sheet sheet, String[] xAxes, int[] yAxes, ExcelCellStyle cellStyle) {
            for (int yAxis : yAxes) {
                Row row = row(sheet, yAxis);
                for (String xAxis : xAxes) {
                    Cell cell = cell(row, xAxis);
                    if (cellStyle != null) {
                        //TODO: 대체하기
//                        cellStyle.applyCellStyle(cell, this.cellStyleMap, this.fontMap);
                    }
                }
            }
        }

        public void applyCellStyle(Cell cell, ExcelCellStyle excelCellStyle) {
            CellStyle cellStyle = cell.getCellStyle();
            if(cellStyle == null) {
                cellStyle = cell.getSheet().getWorkbook().createCellStyle();
            }

        }

        //TODO: row style 은 고민해 보기
        public void applyRowStyle(Row row, ExcelCellStyle excelCellStyle) {
        }

        public Row row(Sheet sheet, int yAxis) {
            int rowNum = ExcelUtils.yAxisToRownum(yAxis);
            Row row = sheet.getRow(rowNum);
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
