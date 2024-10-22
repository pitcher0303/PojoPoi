package com.pojo.poi.core.excel;

import org.apache.poi.util.Units;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class ExcelUtils {
    public static Integer xAxisToCellNum(String xAxis) {
        return (int) xAxis.charAt(0) - 65;
    }

    public static String cellNumToXAxis(int cellNum) {
        return String.valueOf((char) (cellNum + 65));
    }

    public static List<Integer> xAxisToCellNums(String... xAxis) {
        return Arrays.stream(xAxis)
                .map(ExcelUtils::xAxisToCellNum)
                .collect(Collectors.toList());
    }

    public static Integer yAxisToRownum(int yAxis) {return yAxis - 1;}

    public static Integer sumYAxis(int... yAxis) {
        int sum = 0;
        for(int y : yAxis) {
            sum += y;
        }
        return yAxis.length > 1 ? sum - yAxis.length + 1 : sum;
    }

    public static Integer rownumToYAxis(int rownum) {return rownum + 1;}

    public static List<Integer> yAxisToRownums(int... yAxis) {
        return IntStream.of(yAxis)
                .map(ExcelUtils::yAxisToRownum)
                .collect(ArrayList::new, List::add, List::addAll);
    }

    public static List<Integer> AxisFromToNums(List<Integer> axis) {
        if(axis.size() == 1) return axis;
        axis.sort(Integer::compareTo);
        return IntStream.range(axis.getFirst(), axis.getLast() + 1)
                .collect(ArrayList::new, List::add, List::addAll);
    }

    public static List<String> xAxisFromToxAxis(String... xAxis) {
        List<Integer> cellNums = xAxisToCellNums(xAxis);
        return AxisFromToNums(cellNums)
                .stream()
                .map(ExcelUtils::cellNumToXAxis)
                .collect(Collectors.toList());
    }

    public static List<Integer> yAxisFromToxAxis(int... yAxis) {
        List<Integer> yAxisList = yAxisToRownums(yAxis);
        return AxisFromToNums(yAxisList)
                .stream()
                .map(ExcelUtils::rownumToYAxis)
                .collect(Collectors.toList());
    }

    public static boolean isMergedCell(List<Integer> cellnum, List<Integer> rownum) {
        return cellnum.size() > 1 || rownum.size() > 1;
    }

    /**
     * Excel Cell Width 값을 Poi Width 값으로 변환
     *
     * @param excelCellWith excel 상 width 값
     * @return Poi Width 값
     */
    public static int width256(float excelCellWith) {
        return (int) Math.floor((excelCellWith * Units.DEFAULT_CHARACTER_WIDTH + 5) / Units.DEFAULT_CHARACTER_WIDTH * 256);
    }
}
