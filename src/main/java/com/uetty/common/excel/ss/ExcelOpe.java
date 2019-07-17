package com.uetty.common.excel.ss;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

// 暂时没用，计划后续自己实现POI的封装的
@Deprecated()
public class ExcelOpe {

    /**
     * @return 读出的Excel中数据的内容
     */
    public static Map<String, Map<String, List<String>>> getData(String path) throws FileNotFoundException, IOException {
        Map<String, Map<String, List<String>>> table = Maps.newHashMap();
        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        String cellData = null;
        wb = readExcel(path);
        if (wb != null) {
            //遍历所有Sheet
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Map<String, List<String>> sheetMap = Maps.newHashMap();
                sheet = wb.getSheetAt(i);
                //获取最大行数
                int rownum = sheet.getPhysicalNumberOfRows();
                //获取第一行
                Row row0 = sheet.getRow(0);
                //获取最大列数
                int colnum = row0.getPhysicalNumberOfCells();
                //创建初始数据
                for (int colnumIndex = 0; colnumIndex < colnum; colnumIndex++) {
                    String key = getCellFormatValue(row0.getCell(colnumIndex));
                    sheetMap.put(key, Lists.newArrayList());
                }
                //循环存储数据
                for (int rowIndex = 1; rowIndex < rownum; rowIndex++) {
                    row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int colnumIndex = 0; colnumIndex < colnum; colnumIndex++) {
                            List<String> list = sheetMap.get(getCellFormatValue(row0.getCell(colnumIndex)));
                            cellData = getCellFormatValue(row.getCell(colnumIndex));
                            list.add(cellData);
                        }
                    }
                }
                table.put(sheet.getSheetName(), sheetMap);
            }
        }
        return table;
    }

    //读取excel
    public static Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                return wb = new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                return wb = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static String getCellFormatValue(Cell cell) {
        String cellValue = null;
        if (cell != null) {
            int code = cell.getCellType();
            CellType cellType = CellType._NONE;
            for (CellType value : CellType.values()) {
                if (value.getCode() == code) {
                    cellType = value;
                    break;
                }
            }
            //判断cell类型
            switch (cellType) {
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case FORMULA:
                    //判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //转换为日期格式YYYY-mm-dd
                        cellValue = sdf.format(cell.getDateCellValue());
                    } else {
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    public static void writeData(String translateFileName, Map<String, Map<String, List<String>>> copy) {
        String extString = translateFileName.substring(translateFileName.lastIndexOf("."));
        if (".xls".equals(extString)) {
            writeDataXls(translateFileName, copy);
        } else if (".xlsx".equals(extString)) {
            writeDataXlsx(translateFileName, copy);
        }
    }

    /**
     * 输出
     */
    public static void writeDataXlsx(String translateFileName, Map<String, Map<String, List<String>>> copy) {

        try (FileOutputStream out = new FileOutputStream(translateFileName)) {
            //新建文件
            //创建workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            copy.forEach((key, value) -> {
                XSSFSheet sheet = workbook.createSheet(key);
                //添加表头
                XSSFRow row = workbook.getSheet(key).createRow(0);
                AtomicInteger colnumIndex = new AtomicInteger();
                Map<Integer, XSSFRow> map = Maps.newHashMap();
                value.forEach((rowName, colnumList) -> {
                    XSSFCell cell = row.createCell(colnumIndex.get());
                    cell.setCellValue(rowName);
                    //添加该列内容
                    for (int rowIndex = 0; rowIndex < colnumList.size(); rowIndex++) {
                        XSSFRow rows = map.get(rowIndex);
                        if (rows == null) {
                            rows = workbook.getSheet(key).createRow(rowIndex + 1);
                            map.put(rowIndex, rows);
                        }
                        cell = rows.createCell(colnumIndex.get());
                        cell.setCellValue(colnumList.get(rowIndex));
                    }
                    colnumIndex.getAndIncrement();
                });
            });
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 输出
     */
    public static void writeDataXls(String translateFileName, Map<String, Map<String, List<String>>> copy) {

        try (FileOutputStream out = new FileOutputStream(translateFileName);) {
            //新建文件
            //创建workbook
            HSSFWorkbook workbook = new HSSFWorkbook();
            copy.forEach((key, value) -> {
                HSSFSheet sheet = workbook.createSheet(key);
                //添加表头
                HSSFRow row = workbook.getSheet(key).createRow(0);
                int colnumIndex = 0;
                value.forEach((rowName, colnumList) -> {
                    HSSFCell cell = row.createCell(colnumIndex);
                    cell.setCellValue(rowName);
                    //添加该列内容
                    for (int rowIndex = 1; rowIndex < colnumList.size(); rowIndex++) {
                        cell = row.createCell(rowIndex);
                        cell.setCellValue(colnumList.get(rowIndex - 1));
                    }
                });
            });
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
