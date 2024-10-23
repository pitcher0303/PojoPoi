package com.pojo.poi.core.excel.style;

import com.pojo.poi.core.excel.ExcelUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 기본적인 셀 스타일
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellStyle {
    boolean wrapText() default false;

    HorizontalAlignment[] horizontalAlignment() default {};

    VerticalAlignment[] verticalAlignment() default {};

    BorderStyle[] borderTop() default {};

    BorderStyle[] borderRight() default {};

    BorderStyle[] borderBottom() default {};

    BorderStyle[] borderLeft() default {};

    FillPatternType[] fillPattern() default {};

    ExcelColor[] foregroundColor() default {};

    ExcelColor[] backgroundColor() default {};

    ExcelFont[] font() default {};

    //TODO: CellStyle 을 계속 생성하는것 보다는 Model 쪽에서 cellstyle 관리자를 둬서 재활용 하는게 나아보임.
    public static class Creator {
        public static void apply(Cell cell, ExcelCellStyle[] excelCellStyle) {
            if (excelCellStyle == null || excelCellStyle.length < 1) return;
            CellStyle cellStyle = createOrGet(cell);
            ExcelCellStyle target = excelCellStyle[0];
            if (target.wrapText()) {
                cellStyle.setWrapText(true);
            }
            boolean horizontalAlignmentApply = ExcelUtils.isApply(target.horizontalAlignment());
            if (horizontalAlignmentApply) {
                cellStyle.setAlignment(target.horizontalAlignment()[0]);
            }
            boolean verticalAlignmentApply = ExcelUtils.isApply(target.verticalAlignment());
            if (verticalAlignmentApply) {
                cellStyle.setVerticalAlignment(target.verticalAlignment()[0]);
            }
            boolean borderTopApply = ExcelUtils.isApply(target.borderTop());
            if (borderTopApply) {
                cellStyle.setBorderTop(target.borderTop()[0]);
            }
            boolean borderRightApply = ExcelUtils.isApply(target.borderRight());
            if (borderRightApply) {
                cellStyle.setBorderRight(target.borderRight()[0]);
            }
            boolean borderBottomApply = ExcelUtils.isApply(target.borderBottom());
            if (borderBottomApply) {
                cellStyle.setBorderBottom(target.borderBottom()[0]);
            }
            boolean borderLeftApply = ExcelUtils.isApply(target.borderLeft());
            if (borderLeftApply) {
                cellStyle.setBorderLeft(target.borderLeft()[0]);
            }
            boolean fillPatternApply = ExcelUtils.isApply(target.fillPattern());
            if (fillPatternApply) {
                cellStyle.setFillPattern(target.fillPattern()[0]);
            }
            boolean foregroundColorApply = ExcelUtils.isApply(target.foregroundColor());
            if (foregroundColorApply) {
                ExcelColor color = target.foregroundColor()[0];
                ExcelColor.Applier.apply(cellStyle, color);
            }
            boolean backgroundColorApply = ExcelUtils.isApply(target.backgroundColor());
            if (backgroundColorApply) {
                ExcelColor color = target.backgroundColor()[0];
                ExcelColor.Applier.apply(cellStyle, color);
            }
            boolean fontApply = ExcelUtils.isApply(target.font());
            if (fontApply) {
                Font font = cell.getSheet().getWorkbook().createFont();
                ExcelFont.Applier.apply(font, target.font()[0]);
                cellStyle.setFont(font);
            }
            cell.setCellStyle(cellStyle);
        }

        private static CellStyle createOrGet(Cell cell) {
            return cell.getSheet().getWorkbook().createCellStyle();
//            if (cell.getCellStyle() == null) {
//                CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
//                cell.setCellStyle(cellStyle);
//            }
//            return cell.getCellStyle();
        }
    }
}
