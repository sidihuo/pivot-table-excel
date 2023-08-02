package com.sidihuo.pivottableexcel;

import com.sidihuo.pivottable.PivotTableOutput;
import com.sidihuo.pivottable.convert.PivotHelper;
import com.sidihuo.pivottable.exception.PivotTableException;
import com.sidihuo.pivottable.model.output.OutputDataColumn;
import com.sidihuo.pivottable.model.output.OutputDataRow;
import com.sidihuo.pivottable.model.output.OutputHeaderRow;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @Description
 * @Date 2023/7/31 15:24
 * @Created by yanggangjie
 */
@Slf4j
public class PivotExcelWriter {
    public static void write(PivotTableOutput pivotTableOutput,XSSFWorkbook xSSFWorkbook,String sheetName) {
        log.info("begin write PivotExcel {}", VersionLog.JAR_BUILD_TIME);
        try {
            doWrite(pivotTableOutput,xSSFWorkbook,sheetName);
        } catch (Throwable t) {
            log.warn("write PivotExcel failed {}", VersionLog.JAR_BUILD_TIME, t);
            if (t instanceof PivotTableException) {
                throw (PivotTableException) t;
            } else {
                throw new PivotTableException("8000", "write pivot table excel failed", t);
            }
        }
        log.info("end write PivotExcel {}", VersionLog.JAR_BUILD_TIME);
    }


    private static void doWrite(PivotTableOutput pivotTableOutput,XSSFWorkbook xSSFWorkbook,String sheetName) {


        XSSFWorkbook wb = xSSFWorkbook;
        // sheet1
        Sheet sheet1 = wb.createSheet(sheetName);

        //设置样式
//        XSSFCellStyle style = wb.createCellStyle();
//        style.setAlignment(HorizontalAlignment.CENTER); // 设置水平对齐方式为居中对齐
//        style.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中对齐
        //标题头
        //sheet1.createFreezePane(0, 1);//冻结表头

        OutputHeaderRow headerRow = pivotTableOutput.getHeaderRow();
        Map<Integer, List<String>> integerListMap = PivotHelper.buildHeaderCellsMap(headerRow);
        Set<Integer> rowIndexHeaderSet = integerListMap.keySet();
        List<Integer> rowIndexHeaderList = new ArrayList<Integer>(rowIndexHeaderSet);
        Collections.sort(rowIndexHeaderList);
        List<String> rowHeaders = headerRow.getRowHeaders();
        int rowIndex = 0;
        for (Integer integer : rowIndexHeaderList) {
            Row rowTemp = sheet1.createRow(rowIndex++);
            int columnIndex = 0;
            for (String rowHeader : rowHeaders) {
                Cell cell = rowTemp.createCell(columnIndex++);
                cell.setCellValue(rowHeader);
            }
            List<String> strings = integerListMap.get(integer);
            int rowNum = rowTemp.getRowNum();
            int columnMergedFirst = 0;
            int columnMergedLast = 0;
            String columnMergedCurrent = null;
            for (String string : strings) {
                Cell cell = rowTemp.createCell(columnIndex++);
                cell.setCellValue(string);
                if (columnMergedCurrent == null) {
                    columnMergedFirst = cell.getColumnIndex();
                    columnMergedLast = columnMergedFirst;
                    columnMergedCurrent = string;
                } else {
                    if (StringUtils.equals(columnMergedCurrent, string)) {
                        columnMergedLast = cell.getColumnIndex();
                    } else {
                        if (columnMergedLast > columnMergedFirst) {
                            sheet1.addMergedRegion(new CellRangeAddress(rowNum, rowNum, columnMergedFirst, columnMergedLast));
                            columnMergedFirst = cell.getColumnIndex();
                            columnMergedLast = columnMergedFirst;
                            columnMergedCurrent = string;
                        } else {
                            columnMergedFirst = cell.getColumnIndex();
                            columnMergedLast = columnMergedFirst;
                            columnMergedCurrent = string;
                        }
                    }
                }
            }
            if (columnMergedLast > columnMergedFirst) {
                sheet1.addMergedRegion(new CellRangeAddress(rowNum, rowNum, columnMergedFirst, columnMergedLast));
            }
        }

        int rowMergedFirst = 0;
        int rowMergedLast = rowIndexHeaderList.size() - 1;
        int columnMergedCurrent = 0;
        for (String rowHeader : rowHeaders) {
            sheet1.addMergedRegion(new CellRangeAddress(rowMergedFirst, rowMergedLast, columnMergedCurrent, columnMergedCurrent));
            columnMergedCurrent++;
        }


        List<OutputDataRow> dataRows = pivotTableOutput.getDataRows();
        for (OutputDataRow dataRow : dataRows) {
            Row rowTemp = sheet1.createRow(rowIndex++);
            int columnIndex = 0;
            List<String> rowHeaderValues = dataRow.getRowHeaderValues();
            for (String rowHeaderValue : rowHeaderValues) {
                Cell cell = rowTemp.createCell(columnIndex++);
                cell.setCellValue(rowHeaderValue);
            }
            List<OutputDataColumn> columnValues = dataRow.getColumnValues();
            for (OutputDataColumn columnValue : columnValues) {
                Double value = columnValue.getValue();
                Cell cell = rowTemp.createCell(columnIndex++);
                if (value != null) {
                    cell.setCellValue(value.doubleValue());
                }
            }
        }

        int rowMergeBeginIndex = rowIndexHeaderList.size();
        List<String> rowHeaderValuesMergeCurrentValue = null;
        List<Integer> rowHeaderValuesMergeCurrentRowFirst = null;
        for (OutputDataRow dataRow : dataRows) {
            List<String> rowHeaderValues = dataRow.getRowHeaderValues();
            if (rowHeaderValuesMergeCurrentValue == null) {
                rowHeaderValuesMergeCurrentValue = rowHeaderValues;
                rowHeaderValuesMergeCurrentRowFirst = new ArrayList<Integer>();
                for (String s : rowHeaderValuesMergeCurrentValue) {
                    rowHeaderValuesMergeCurrentRowFirst.add(rowMergeBeginIndex);
                }
            } else {
                int size = rowHeaderValuesMergeCurrentValue.size();
                for (int columnIndex = size - 1; columnIndex >= 0; columnIndex--) {
                    String last = rowHeaderValuesMergeCurrentValue.get(columnIndex);
                    String current = rowHeaderValues.get(columnIndex);
                    if (StringUtils.equals(last, current)) {
                    } else {
                        rowHeaderValuesMergeCurrentValue.set(columnIndex, current);
                        int mergeRowFirst = rowHeaderValuesMergeCurrentRowFirst.get(columnIndex);
                        int mergeRowLast = rowMergeBeginIndex - 1;
                        if (mergeRowLast > mergeRowFirst) {
                            sheet1.addMergedRegion(new CellRangeAddress(mergeRowFirst, mergeRowLast, columnIndex, columnIndex));
                        }
                        rowHeaderValuesMergeCurrentRowFirst.set(columnIndex, rowMergeBeginIndex);
                    }
                }

            }
            rowMergeBeginIndex++;
        }

        OutputDataRow outputDataRowLastLine = dataRows.get(dataRows.size() - 1);
        List<String> rowHeaderValues = outputDataRowLastLine.getRowHeaderValues();
        int size = rowHeaderValuesMergeCurrentValue.size();
        for (int columnIndex = size - 1; columnIndex >= 0; columnIndex--) {
            String last = rowHeaderValuesMergeCurrentValue.get(columnIndex);
            String current = rowHeaderValues.get(columnIndex);
            if (StringUtils.equals(last, current)) {
                int mergeRowFirst = rowHeaderValuesMergeCurrentRowFirst.get(columnIndex);
                int mergeRowLast = rowMergeBeginIndex - 1;
                if (mergeRowLast > mergeRowFirst) {
                    sheet1.addMergedRegion(new CellRangeAddress(mergeRowFirst, mergeRowLast, columnIndex, columnIndex));
                }
            }
        }

    }
}
