package com.pojo.poi.core.excel.model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Map;

public enum ExcelCellStyle {
    NONE,
    MODEL_META_VALUE,
    INFO_CELL,
    CENTER_INFO_CELL,
    HEADER_CELL1,
    HEADER_CELL2,
    HEADER_CELL3;

    public void applyCellStyle(Cell cell, Map<ExcelCellStyle, CellStyle> cellStyleMap, Map<ExcelCellStyle, Font> fontMap) {
        CellStyle cellStyle = null;

        if(!cellStyleMap.containsKey(this)) {
            cellStyle = cell.getSheet().getWorkbook().createCellStyle();
            cellStyleMap.put(this, cellStyle);
        }

        cellStyle = cellStyleMap.get(this);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        if(!fontMap.containsKey(this)) {
            fontMap.put(this, new XSSFFont());
        }

        Font font = fontMap.get(this);
        font.setFontName("");
        font.setFontHeightInPoints((short) 11);
        switch(this) {
            case NONE -> {
                break;
            }
            //파란 굵은 글씨
            case MODEL_META_VALUE -> {
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
                byte[] rgb = {(byte) 0, (byte) 0, (byte) 128};
                ((XSSFFont)font).setColor(new XSSFColor(rgb));
                font.setBold(true);
            }
            //얇은 테두리
            case INFO_CELL -> {
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
            }
            //얇은 테두리 | 가운데 정렬
            case CENTER_INFO_CELL -> {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
            }
            //왼쪽, 위쪽, 위 얇은 테두리 | 아래 굵은 테두리 | 굵은 글씨 | 회색
            case HEADER_CELL1 -> {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                byte[] rgb = {(byte) 192, (byte) 192, (byte) 192};
                ((XSSFCellStyle)cellStyle).setFillBackgroundColor(new XSSFColor(rgb));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                font.setBold(true);
            }
            //왼쪽, 오른쪽 얇은 테두리 | 위, 아래 굵은 테두리 | 굵은 글씨 | 가운데 정렬 | 회색
            case HEADER_CELL2 -> {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                byte[] rgb = {(byte) 192, (byte) 192, (byte) 192};
                ((XSSFCellStyle)cellStyle).setFillBackgroundColor(new XSSFColor(rgb));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.MEDIUM);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.MEDIUM);
                font.setBold(true);
            }
            //얇은 테두리 | 굵은 글씨 | 가운데 정렬 | 주황색
            case HEADER_CELL3 -> {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                byte[] rgb = {(byte) 255, (byte) 204, (byte) 153};
                ((XSSFCellStyle)cellStyle).setFillBackgroundColor(new XSSFColor(rgb));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                font.setBold(true);
            }
        }
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }
}
