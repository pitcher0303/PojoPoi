package com.pojo.poi.core.excel.style;

import com.pojo.poi.core.excel.ExcelUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

public class ExcelStyleManager {
    private final Map<Integer, CellStyle> cellStyleCache;
    private final Map<Integer, Font> fonts;
    private final Map<String, CellStyle> rowStylesCache;
    private final Workbook workbook;

    public ExcelStyleManager(Workbook workbook) {
        this.cellStyleCache = new HashMap<>();
        this.fonts = new HashMap<>();
        this.rowStylesCache = new HashMap<>();
        this.workbook = workbook;
    }

    public CellStyle getCellStyle(ExcelCellStyle excelCellStyle) {
        CellStyle cellStyle = this.cellStyleCache.get(excelCellStyle.hashCode());
        if (cellStyle == null) {
            cellStyle = this.workbook.createCellStyle();
            this.cellStyleCache.put(excelCellStyle.hashCode(), cellStyle);
        }
        return cellStyle;
    }

    public Font getFont(ExcelFont excelFont) {
        Font font = this.fonts.get(excelFont.hashCode());
        if (font == null) {
            font = this.workbook.createFont();
            this.fonts.put(excelFont.hashCode(), font);
        }
        return font;
    }

    public void applyCellStyle(Cell cell, ExcelCellStyle[] excelCellStyles) {
        if (excelCellStyles == null || excelCellStyles.length < 1) return;
        ExcelCellStyle excelCellStyle = excelCellStyles[0];
        if (isDefault(excelCellStyle)) return;
        CellStyle cellStyle = getCellStyle(excelCellStyle);
        applyCellStyle(cellStyle, excelCellStyle);
        cell.setCellStyle(cellStyle);
    }

    private void applyCellStyle(CellStyle cellStyle, ExcelCellStyle excelCellStyle) {
        if (excelCellStyle.wrapText()) {
            cellStyle.setWrapText(true);
        }
        boolean horizontalAlignmentApply = ExcelUtils.isApply(excelCellStyle.horizontalAlignment());
        if (horizontalAlignmentApply) {
            cellStyle.setAlignment(excelCellStyle.horizontalAlignment()[0]);
        }
        boolean verticalAlignmentApply = ExcelUtils.isApply(excelCellStyle.verticalAlignment());
        if (verticalAlignmentApply) {
            cellStyle.setVerticalAlignment(excelCellStyle.verticalAlignment()[0]);
        }
        boolean borderTopApply = ExcelUtils.isApply(excelCellStyle.borderTop());
        if (borderTopApply) {
            cellStyle.setBorderTop(excelCellStyle.borderTop()[0]);
        }
        boolean borderRightApply = ExcelUtils.isApply(excelCellStyle.borderRight());
        if (borderRightApply) {
            cellStyle.setBorderRight(excelCellStyle.borderRight()[0]);
        }
        boolean borderBottomApply = ExcelUtils.isApply(excelCellStyle.borderBottom());
        if (borderBottomApply) {
            cellStyle.setBorderBottom(excelCellStyle.borderBottom()[0]);
        }
        boolean borderLeftApply = ExcelUtils.isApply(excelCellStyle.borderLeft());
        if (borderLeftApply) {
            cellStyle.setBorderLeft(excelCellStyle.borderLeft()[0]);
        }
        boolean fillPatternApply = ExcelUtils.isApply(excelCellStyle.fillPattern());
        if (fillPatternApply) {
            cellStyle.setFillPattern(excelCellStyle.fillPattern()[0]);
        }
        boolean foregroundColorApply = ExcelUtils.isApply(excelCellStyle.foregroundColor());
        if (foregroundColorApply) {
            ExcelColor color = excelCellStyle.foregroundColor()[0];
            ExcelColor.Applier.apply(cellStyle, color);
            if (!fillPatternApply) {
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }
        boolean backgroundColorApply = ExcelUtils.isApply(excelCellStyle.backgroundColor());
        if (backgroundColorApply) {
            ExcelColor color = excelCellStyle.backgroundColor()[0];
            ExcelColor.Applier.apply(cellStyle, color);
        }
        boolean fontApply = ExcelUtils.isApply(excelCellStyle.font());
        if (fontApply) {
            Font font = getFont(excelCellStyle.font()[0]);
            ExcelFont.Applier.apply(font, excelCellStyle.font()[0]);
            cellStyle.setFont(font);
        }
    }

    private boolean isDefault(ExcelCellStyle excelCellStyle) {
        boolean wrapText = excelCellStyle.wrapText();
        boolean horizontalAlignmentApply = ExcelUtils.isApply(excelCellStyle.horizontalAlignment());
        boolean verticalAlignmentApply = ExcelUtils.isApply(excelCellStyle.verticalAlignment());
        boolean borderTopApply = ExcelUtils.isApply(excelCellStyle.borderTop());
        boolean borderRightApply = ExcelUtils.isApply(excelCellStyle.borderRight());
        boolean borderBottomApply = ExcelUtils.isApply(excelCellStyle.borderBottom());
        boolean borderLeftApply = ExcelUtils.isApply(excelCellStyle.borderLeft());
        boolean fillPatternApply = ExcelUtils.isApply(excelCellStyle.fillPattern());
        boolean foregroundColorApply = ExcelUtils.isApply(excelCellStyle.foregroundColor());
        boolean backgroundColorApply = ExcelUtils.isApply(excelCellStyle.backgroundColor());
        boolean fontApply = ExcelUtils.isApply(excelCellStyle.font());
        return !wrapText && !horizontalAlignmentApply && !verticalAlignmentApply
                && !borderTopApply && !borderRightApply && !borderBottomApply
                && !borderLeftApply && !fillPatternApply && !foregroundColorApply
                && !backgroundColorApply && !fontApply;
    }

    public void applyRowStyle(List<Row> fromToRows, ExcelRowStyle[] excelRowStyles) {
        if (excelRowStyles.length < 1) return;
        ExcelRowStyle excelRowStyle = excelRowStyles[0];
        boolean isBorderLeft = ExcelUtils.isApply(excelRowStyle.borderLeft());
        boolean isBorderRight = ExcelUtils.isApply(excelRowStyle.borderRight());
        boolean isBorderTop = ExcelUtils.isApply(excelRowStyle.borderTop());
        boolean isBorderBottom = ExcelUtils.isApply(excelRowStyle.borderBottom());
        for (int i = 0; i < fromToRows.size(); i++) {
            Row row = fromToRows.get(i);
            if (isBorderLeft) {
                Cell cell = row.getCell(row.getFirstCellNum());
                CellStyle cellStyle = getRowCellStyle(cell, RowBorders.LEFT, excelRowStyle.borderLeft()[0]);
                cell.setCellStyle(cellStyle);
            }
            if (isBorderRight) {
                Cell cell = row.getCell(row.getLastCellNum() - 1);
                CellStyle cellStyle = getRowCellStyle(cell, RowBorders.RIGHT, excelRowStyle.borderRight()[0]);
                cell.setCellStyle(cellStyle);
            }
            if (i == 0) {
                if (isBorderTop) {
                    for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
//                        Cell cell = row.getCell(cellNum);
                        Cell cell = ExcelUtils.cell(row, ExcelUtils.cellNumToXAxis(cellNum));
                        CellStyle cellStyle = getRowCellStyle(cell, RowBorders.TOP, excelRowStyle.borderTop()[0]);
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
            if (i == fromToRows.size() - 1) {
                if (isBorderBottom) {
                    for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        CellStyle cellStyle = getRowCellStyle(cell, RowBorders.BOTTOM, excelRowStyle.borderBottom()[0]);
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
        }
    }

    public CellStyle getRowCellStyle(Cell cell, RowBorders type, BorderStyle borderStyle) {
        CellStyle cellStyle = cell.getCellStyle();
        String key = type.generateRowStyleId(cellStyle.hashCode());
        if (!rowStylesCache.containsKey(key)) {
            CellStyle temp = cell.getSheet().getWorkbook().createCellStyle();
            temp.cloneStyleFrom(cellStyle);
            switch (type) {
                case TOP -> temp.setBorderTop(borderStyle);
                case RIGHT -> temp.setBorderRight(borderStyle);
                case BOTTOM -> temp.setBorderBottom(borderStyle);
                case LEFT -> temp.setBorderLeft(borderStyle);
            }
            String newKey = type.generateRowStyleId(temp.hashCode());
            rowStylesCache.put(newKey, temp);
            cellStyle = temp;
        }
        return cellStyle;
    }

    public enum RowBorders {
        TOP((hascode) -> "TOP" + hascode),
        RIGHT((hascode) -> "RIGHT" + hascode),
        BOTTOM((hascode) -> "BOTTOM" + hascode),
        LEFT((hascode) -> "LEFT" + hascode);

        private final Function<Integer, String> idGenerator;

        RowBorders(Function<Integer, String> idGenerator) {
            this.idGenerator = idGenerator;
        }

        public String generateRowStyleId(Integer hascode) {
            return this.idGenerator.apply(hascode);
        }
    }
}
