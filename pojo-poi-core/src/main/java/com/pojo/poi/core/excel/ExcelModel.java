package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.*;
import com.pojo.poi.core.excel.style.ExcelCellStyle;
import com.pojo.poi.core.excel.style.ExcelRowStyle;
import com.pojo.poi.core.excel.style.ExcelStyleManager;
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
    private final ExcelStyleManager excelStyleManager;

    public ExcelModel(XSSFWorkbook workbook, String originFileName, String fileName, int sheetIndex) {
        this.workbook = workbook;
        this.originFileName = originFileName;
        this.fileName = fileName;
        this.sheetIndex = sheetIndex;
        this.excelStyleManager = new ExcelStyleManager(workbook);
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

    public ExcelModel addExcelDatas(final String sheetName, List<ExcelData> excelDatas) {
        return addExcelDatas(sheetName, excelDatas, null);
    }

    public ExcelModel addExcelDatas(final String sheetName, List<ExcelData> excelDatas, float[] cellWidths) {
        Optional<ExcelSheetModel> sheetModel = this.sheets.stream()
                .filter(excelSheetModel -> excelSheetModel.sheetName.equals(sheetName))
                .findAny();
        sheetModel.ifPresentOrElse(excelSheetModel -> excelSheetModel.excelDatas.addAll(excelDatas), () -> {
            ExcelSheetModel excelSheetModel = new ExcelSheetModel(this.workbook, sheetName, cellWidths, this.excelStyleManager);
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
        private final ExcelStyleManager excelStyleManager;

        public ExcelSheetModel(Workbook workbook,
                               String sheetName,
                               float[] cellWidths,
                               ExcelStyleManager excelStyleManager) {
            this.sheet = this.cresheet(workbook, sheetName);
            this.sheetName = sheetName;
            this.excelDatas = new ArrayList<>();
            this.setColumnWidths(cellWidths);
            this.excelStyleManager = excelStyleManager;
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
