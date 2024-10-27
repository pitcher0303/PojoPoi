package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.style.ExcelStyleManager;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

@Getter
public class ExcelWriter {
    private static final String DEFAULT_SHEET_NAME = "sheet";
    private final XSSFWorkbook workbook;
    private final String originFileName;
    private final List<ExcelSheetWriter> sheets = new ArrayList<>();
    private String fileName;
    private int sheetIndex = 1;
    private final ExcelStyleManager excelStyleManager;

    public ExcelWriter(XSSFWorkbook workbook, String originFileName, String fileName, int sheetIndex) {
        this.workbook = workbook;
        this.originFileName = originFileName;
        this.fileName = fileName;
        this.sheetIndex = sheetIndex;
        this.excelStyleManager = new ExcelStyleManager(workbook);
    }

    public static Builder builder(String fileName) {
        if (!fileName.endsWith(".xlsx")) fileName = fileName.concat(".xlsx");
        return Builder.builder(fileName);
    }

    public ExcelWriter fileName(String fileName) {
        if (!fileName.endsWith(".xlsx")) fileName = fileName.concat(".xlsx");
        this.fileName = fileName;
        return this;
    }

    public ExcelWriter addExcelDatas(List<ExcelData> excelDatas) {
        String sheetName = DEFAULT_SHEET_NAME + sheetIndex++;
        return addExcelDatas(sheetName, excelDatas, null);
    }

    public ExcelWriter addExcelDatas(final String sheetName, List<ExcelData> excelDatas) {
        return addExcelDatas(sheetName, excelDatas, null);
    }

    public ExcelWriter addExcelDatas(final String sheetName, List<ExcelData> excelDatas, float[] cellWidths) {
        Optional<ExcelSheetWriter> sheetModel = this.sheets.stream()
                .filter(excelSheetWriter -> excelSheetWriter.sheetName.equals(sheetName))
                .findAny();
        sheetModel.ifPresentOrElse(excelSheetModel -> excelSheetModel.excelDatas.addAll(excelDatas), () -> {
            ExcelSheetWriter excelSheetModel = new ExcelSheetWriter(this.workbook, sheetName, cellWidths, this.excelStyleManager);
            excelSheetModel.excelDatas.addAll(excelDatas);
            this.sheets.add(excelSheetModel);
        });
        return this;
    }

    public ExcelWriter writeAll() {
        this.sheets.forEach(ExcelSheetWriter::write);
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

    private static class ExcelSheetWriter {
        private final float[] defaultCellWiths = {};
        private final Sheet sheet;
        private final String sheetName;
        private final List<ExcelData> excelDatas;
        private final ExcelStyleManager excelStyleManager;

        public ExcelSheetWriter(Workbook workbook,
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
            //TODO: 추후 sxssfworkbook writer 대비
            excelDatas.forEach(excelData -> XssfExcelWriter.writeExcelData(sheet, excelStyleManager, excelData, startYAxis));
        }

    }

    public static class Builder {
        private XSSFWorkbook workbook;
        private String originFileName;
        private String fileName;

        private static Builder builder(String excelFileName) {
            Builder builder = new Builder();
            builder.originFileName = excelFileName;
            builder.workbook = new XSSFWorkbook();
            builder.fileName = excelFileName;
            return builder;
        }

        public Builder setWorkbook(XSSFWorkbook workbook) {
            this.workbook = workbook;
            return this;
        }

        public Builder setOriginFileName(String originFileName) {
            this.originFileName = originFileName;
            return this;
        }

        public Builder setFileName(String fileName) {
            this.fileName = fileName;
            return this;
        }

        public ExcelWriter build() {
            return new ExcelWriter(this.workbook, this.originFileName, this.fileName, 1);
        }
    }
}
