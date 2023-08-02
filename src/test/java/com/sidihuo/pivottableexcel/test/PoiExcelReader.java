package com.sidihuo.pivottableexcel.test;

import com.sidihuo.pivottable.model.input.InputDataColumn;
import com.sidihuo.pivottable.model.input.InputDataColumnHeader;
import com.sidihuo.pivottable.model.input.InputDataRow;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description
 * @Date 2023/8/2 15:31
 * @Created by yanggangjie
 */
public class PoiExcelReader {

    private List<InputDataRow> rows;
    private List<InputDataColumnHeader> headers;
    private int rowIndex = 0;

    public void setRows(List<InputDataRow> rows) {
        this.rows = rows;
    }

    public void setHeaders(List<InputDataColumnHeader> headers) {
        this.headers = headers;
    }

    /**
     * 读取excel内容
     * <p>
     * 用户模式下：
     * 弊端：对于少量的数据可以，单数对于大量的数据，会造成内存占据过大，有时候会造成内存溢出
     * 建议修改成事件模式
     */
    public  void redExcel(String filePath) throws Exception {
        File file = new File(filePath);
        if (!file.exists()){
            throw new Exception("文件不存在!");
        }
        InputStream in = new FileInputStream(file);
        // 读取整个Excel
        XSSFWorkbook sheets = new XSSFWorkbook(in);
        // 获取第一个表单Sheet
        XSSFSheet sheetAt = sheets.getSheetAt(1);
        ArrayList<Map<String, String>> list = new ArrayList<Map<String, String>>();
        Map<Integer,String> headerIndexMap=new HashMap<Integer, String>();

        //默认第一行为标题行，i = 0
        XSSFRow titleRow = sheetAt.getRow(0);
        for (int index = 0; index < titleRow.getPhysicalNumberOfCells(); index++){
            InputDataColumnHeader header=new InputDataColumnHeader();
            XSSFCell titleCell = titleRow.getCell(index);
            header.setName(titleCell.getStringCellValue());
            headers.add(header);
            headerIndexMap.put(index,header.getName());
        }

        // 循环获取每一行数据
        for (int i = 1; i < sheetAt.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = sheetAt.getRow(i);
            //LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
            InputDataRow rowData=new InputDataRow();
            rows.add(rowData);
            List<InputDataColumn> columns=new ArrayList<InputDataColumn>();
            rowData.setColumns(columns);
            // 读取每一格内容
            for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
                XSSFCell cell = row.getCell(index);
                InputDataColumn c=new InputDataColumn();
                columns.add(c);
                c.setIndex(index);
                c.setHeader(headerIndexMap.get(index));
                if(cell!=null){
                    CellType cellType = cell.getCellType();
                    if(CellType.NUMERIC==cellType){
                        c.setValue(cell.getNumericCellValue());
                    }else {
                        c.setValue(cell.getStringCellValue());

                    }
                }

//                XSSFCell titleCell = titleRow.getCell(index);
//                // cell.setCellType(XSSFCell.CELL_TYPE_STRING); 过期，使用下面替换
//                cell.setCellType(CellType.STRING);
//                if (cell.getStringCellValue().equals("")) {
//                    continue;
//                }
//                map.put(getString(titleCell), getString(cell));
            }
//            if (map.isEmpty()) {
//                continue;
//            }
//            list.add(map);
        }
    }

    /**
     * 把单元格的内容转为字符串
     *
     * @param xssfCell 单元格
     * @return String
     */
    public static String getString(XSSFCell xssfCell) {
        if (xssfCell == null) {
            return "";
        }
        if (xssfCell.getCellTypeEnum() == CellType.NUMERIC) {
            return String.valueOf(xssfCell.getNumericCellValue());
        } else if (xssfCell.getCellTypeEnum() == CellType.BOOLEAN) {
            return String.valueOf(xssfCell.getBooleanCellValue());
        } else {
            return xssfCell.getStringCellValue();
        }
    }
}
