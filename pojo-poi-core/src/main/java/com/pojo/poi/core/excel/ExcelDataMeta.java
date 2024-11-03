package com.pojo.poi.core.excel;

import com.pojo.poi.core.excel.annotation.CellMeta;
import com.pojo.poi.core.excel.annotation.MetaOrder;
import com.pojo.poi.core.excel.annotation.RowMeta;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;

@Getter
@Setter
public class ExcelDataMeta {
    private boolean isCellMeta = false;
    private boolean isRowMeta = false;
    private String startXAxis;
    private String endXAxis;
    private int startYAxis;
    private int endYAxis;
    private Field field;
    private int hashcode;
    private MetaOrder metaOrder;

    public ExcelDataMeta(final Field field) {
        this.isCellMeta = field.isAnnotationPresent(CellMeta.class);
        this.isRowMeta = field.isAnnotationPresent(RowMeta.class);
        this.field = field;
        if (this.isCellMeta) {
            CellMeta cellMeta = field.getAnnotation(CellMeta.class);
            this.startXAxis = cellMeta.xAxis()[0];
            this.endXAxis = cellMeta.xAxis()[cellMeta.xAxis().length - 1];
            this.startYAxis = cellMeta.yAxis()[0];
            this.endYAxis = cellMeta.yAxis()[cellMeta.yAxis().length - 1];
            this.hashcode = cellMeta.hashCode();
            this.metaOrder = cellMeta.metaOrder();
        }
        if (this.isRowMeta) {
            RowMeta rowMeta = field.getAnnotation(RowMeta.class);
            this.startXAxis = rowMeta.xAxis()[0];
            this.endXAxis = rowMeta.xAxis()[rowMeta.xAxis().length - 1];
            this.startYAxis = rowMeta.startYAxis();
            this.endYAxis = rowMeta.endYAxis();
            this.hashcode = rowMeta.hashCode();
            this.metaOrder = rowMeta.metaOrder();
        }
    }
}
