package com.sidihuo.pivottableexcel.test;

//import com.alibaba.excel.context.AnalysisContext;
//import com.alibaba.excel.event.AnalysisEventListener;
//import com.sidihuo.pivottable.model.input.InputDataColumn;
//import com.sidihuo.pivottable.model.input.InputDataColumnHeader;
//import com.sidihuo.pivottable.model.input.InputDataRow;
//import lombok.Data;
//import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
//
//@Data
//@Slf4j
//public class ExcelImportListener extends AnalysisEventListener<Map<Integer, String>> {
//
//    private List<InputDataRow> rows;
//    private List<InputDataColumnHeader> headers;
//
//    private int rowIndex = 0;
//
//    public List<InputDataRow> getRows() {
//        return rows;
//    }
//
//    public void setRows(List<InputDataRow> rows) {
//        this.rows = rows;
//    }
//
//    public List<InputDataColumnHeader> getHeaders() {
//        return headers;
//    }
//
//    public void setHeaders(List<InputDataColumnHeader> headers) {
//        this.headers = headers;
//    }
//
//    /**
//     * 重写invokeHeadMap方法，获去表头，如果有需要获取第一行表头就重写这个方法，不需要则不需要重写
//     *
//     * @param headMap Excel每行解析的数据为Map<Integer, String>类型，Integer是Excel的列索引,String为Excel的单元格值
//     * @param context context能获取一些东西，比如context.readRowHolder().getRowIndex()为Excel的行索引，表头的行索引为0，0之后的都解析成数据
//     */
//    @Override
//    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
//        Iterator<Map.Entry<Integer, String>> iterator = headMap.entrySet().iterator();
//        while (iterator.hasNext()) {
//            Map.Entry<Integer, String> next = iterator.next();
//            InputDataColumnHeader header = new InputDataColumnHeader();
//            //header.setIndex(next.getKey());
//            header.setName(next.getValue());
//            headers.add(header);
//        }
//    }
//
//    /**
//     * 重写invoke方法获得除Excel第一行表头之后的数据，
//     * 如果Excel第二行也是表头，那么也会解析到这里，如果不需要就通过判断context.readRowHolder().getRowIndex()跳过
//     *
//     * @param data    除了第一行表头外，数据都会解析到这个方法
//     * @param context 和上面解释一样
//     */
//    public void invoke(Map<Integer, String> data, AnalysisContext context) {
//        InputDataRow row = new InputDataRow();
//        //row.setIndex(rowIndex++);
//        List<InputDataColumn> columns = new ArrayList<InputDataColumn>();
//        row.setColumns(columns);
//        Iterator<Map.Entry<Integer, String>> iterator = data.entrySet().iterator();
//        while (iterator.hasNext()) {
//            Map.Entry<Integer, String> next = iterator.next();
//            InputDataColumn column = new InputDataColumn();
//            column.setIndex(next.getKey());
//            column.setValue(next.getValue());
//            columns.add(column);
//        }
//        rows.add(row);
//    }
//
//    /**
//     * 解析到最后会进入这个方法，需要重写这个doAfterAllAnalysed方法，然后里面调用自己定义好保存方法
//     *
//     * @param context
//     */
//    public void doAfterAllAnalysed(AnalysisContext context) {
//    }
//
//
//}
