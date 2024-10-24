package com.pojo.poi.core.excel.style;

import com.pojo.poi.core.excel.ExcelUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

public class ExcelStyleManager {
    private final Map<Integer, CellStyle> styles;
    private final Map<Integer, Font> fonts;
    private final Workbook workbook;

    public ExcelStyleManager(Workbook workbook) {
        this.styles = new HashMap<>();
        this.fonts = new HashMap<>();
        this.workbook = workbook;
    }

    public CellStyle getCellStyle(ExcelCellStyle excelCellStyle) {
        CellStyle cellStyle = this.styles.get(excelCellStyle.hashCode());
        if (cellStyle == null) {
            cellStyle = this.workbook.createCellStyle();
            this.styles.put(excelCellStyle.hashCode(), cellStyle);
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
}
