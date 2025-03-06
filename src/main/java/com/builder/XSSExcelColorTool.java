package com.joolun.common.freedomexcel.builder;

import com.joolun.common.freedomexcel.entity.Column;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * excel处理工具
 * 功能: 动态生成单级、多级Excel表头，并根据数据长度调整列宽，支持颜色设置
 */
@Data
public class XSSExcelColorTool<T> {

    private XSSFWorkbook workbook;
    private String title;
    private int colWidth = 20;
    private int rowHeight = 20;
    private XSSFCellStyle styleHead;
    private Map<String, XSSFCellStyle> styleBodyMap = new HashMap<>(); // 缓存不同颜色的样式
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    private static final int DEFAULT_WIDTH = 12;
    private static final int WIDTH_FACTOR = 256;
    private static final int MAX_WIDTH = 255 * WIDTH_FACTOR;
    private Map<Integer, Integer> headerWidths = new HashMap<>();
    private Map<Integer, Integer> dataWidths = new HashMap<>();

    // 构造函数
    public XSSExcelColorTool() {
        this.title = "sheet1";
        this.workbook = new XSSFWorkbook();
        initStyles();
    }

    public XSSExcelColorTool(String title) {
        this.title = title;
        this.workbook = new XSSFWorkbook();
        initStyles();
    }

    public XSSExcelColorTool(String title, int colWidth, int rowHeight) {
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new XSSFWorkbook();
        initStyles();
    }

    public XSSExcelColorTool(String title, int colWidth, int rowHeight, String dateFormat) {
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new XSSFWorkbook();
        this.sdf = new SimpleDateFormat(dateFormat);
        initStyles();
    }

    public XSSExcelColorTool(String title, int colWidth, int rowHeight, int flag) {
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.title = title;
        this.workbook = new XSSFWorkbook();
        initStyles(flag);
    }

    // 初始化样式
    private void initStyles() {
        initStyles(0);
    }

    private void initStyles(int styleFlag) {
        // 表头样式
        this.styleHead = workbook.createCellStyle();
        this.styleHead.setAlignment(HorizontalAlignment.CENTER);
        this.styleHead.setVerticalAlignment(VerticalAlignment.CENTER);

        // 初始化默认主体样式
        XSSFCellStyle defaultStyleBody = workbook.createCellStyle();
        defaultStyleBody.setAlignment(HorizontalAlignment.CENTER);
        defaultStyleBody.setVerticalAlignment(VerticalAlignment.CENTER);
        defaultStyleBody.setBorderRight(BorderStyle.THIN);
        defaultStyleBody.setBorderBottom(BorderStyle.THIN);
        defaultStyleBody.setFillForegroundColor(IndexedColors.WHITE1.getIndex());
        defaultStyleBody.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleBodyMap.put("default", defaultStyleBody);

        switch (styleFlag) {
            case 1:
                XSSFCellStyle leftStyle = workbook.createCellStyle();
                leftStyle.cloneStyleFrom(defaultStyleBody);
                leftStyle.setAlignment(HorizontalAlignment.LEFT);
                styleBodyMap.put("left", leftStyle);
                break;
            case 2:
                this.styleHead.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
                this.styleHead.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                Font font = workbook.createFont();
                font.setBold(true);
                font.setColor(IndexedColors.WHITE.getIndex());
                this.styleHead.setFont(font);
                break;
        }

        // 初始化颜色样式
        createColorStyle("red", IndexedColors.RED1);
        createColorStyle("green", IndexedColors.GREEN);
        createColorStyle("blue", IndexedColors.BLUE);
    }

    // 创建带颜色的样式
    private void createColorStyle(String colorName, IndexedColors color) {
        XSSFCellStyle colorStyle = workbook.createCellStyle();
        colorStyle.cloneStyleFrom(styleBodyMap.get("default"));
        colorStyle.setFillForegroundColor(color.getIndex());
        colorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleBodyMap.put(colorName, colorStyle);
    }

    // 获取合适的样式
    private XSSFCellStyle getStyleBody(String color) {
        return styleBodyMap.getOrDefault(color != null ? color.toLowerCase() : "default", styleBodyMap.get("default"));
    }

    // 导出方法
    public void exportExcel(List<Column> headerCellList, List<T> rowList, String filePath, boolean flag) throws Exception {
        splitDataToSheets(rowList, headerCellList, flag, false);
        save(workbook, filePath);
    }

    public void exportExcel(List<Column> headerCellList, List<T> rowList, String filePath, boolean flag, boolean rowFlag) throws Exception {
        splitDataToSheets(rowList, headerCellList, flag, rowFlag);
        save(workbook, filePath);
    }

    public XSSFWorkbook exportWorkbook(List<Column> listTpamscolumn, List<T> datas, boolean flag) throws Exception {
        splitDataToSheets(datas, listTpamscolumn, flag, false);
        return workbook;
    }

    public InputStream exportExcel(List<Column> headerCellList, List<T> datas, boolean flag, boolean rowFlag) throws Exception {
        splitDataToSheets(datas, headerCellList, flag, rowFlag);
        return save(workbook);
    }

    // 写入Sheet
    private void writeSheet(XSSFSheet sheet, List<T> data, List<Column> headerCellList, boolean flag, boolean rowFlag) throws Exception {
        sheet.setDefaultRowHeightInPoints(rowHeight);
        sheet = createHead(sheet, headerCellList.get(0).getTotalRow(), headerCellList.get(0).getTotalCol());
        createHead(headerCellList, sheet, 0);

        if (flag) {
            writeSheetContent(headerCellList, data, sheet, headerCellList.get(0).getTotalRow(), rowFlag);
        }

        // 设置列宽
        for (int col = 0; col < headerCellList.get(0).getTotalCol(); col++) {
            int headerWidth = headerWidths.getOrDefault(col, DEFAULT_WIDTH);
            int dataWidth = dataWidths.getOrDefault(col, DEFAULT_WIDTH);
            int finalWidth = Math.max(headerWidth, dataWidth);
            sheet.setColumnWidth(col, Math.min(finalWidth * WIDTH_FACTOR, MAX_WIDTH));
        }
    }

    // 分割Sheet
    private void splitDataToSheets(List<T> data, List<Column> headerCellList, boolean flag, boolean rowFlag) throws Exception {
        int dataCount = data.size();
        int maxRows = 65535;
        int pieces = dataCount / maxRows + (dataCount % maxRows > 0 ? 1 : 0);

        for (int i = 0; i < pieces; i++) {
            XSSFSheet sheet = workbook.createSheet(title + (i + 1));
            int start = i * maxRows;
            int end = Math.min((i + 1) * maxRows, dataCount);
            List<T> subList = data.subList(start, end);
            writeSheet(sheet, subList, headerCellList, flag, rowFlag);
        }
    }

    // 写入内容
    private void writeSheetContent(List<Column> headerCellList, List<T> datas, XSSFSheet sheet, int rowIndex, boolean rowFlag) throws Exception {
        List<Column> listCol = new ArrayList<>();
        getColumnList(headerCellList, listCol);
        for (int i = 0, index = rowIndex; i < datas.size(); i++, index++) {
            XSSFRow row = sheet.createRow(index);
            for (int j = 0; j < listCol.size(); j++) {
                createCol(row, listCol.get(j), datas.get(i));
            }
        }
    }

    // 创建单元格
    public void createCol(XSSFRow row, Column tpamscolumn, T v) throws Exception {
        XSSFCell cell = row.createCell(tpamscolumn.getCol());
        Object value = null;
        String color = null;

        if (v instanceof Map) {
            Map<?, ?> m = (Map<?, ?>) v;
            value = m.get(tpamscolumn.getFieldName());
            if (m.get(tpamscolumn.getFieldName() + "_color") != null) {
                color = m.get(tpamscolumn.getFieldName() + "_color").toString();
            }
        } else {
            Class<?> cls = v.getClass();
            for (Field f : cls.getDeclaredFields()) {
                f.setAccessible(true);
                if (tpamscolumn.getFieldName().equals(f.getName()) && !tpamscolumn.isHasChildren()) {
                    value = f.get(v);
                    if (value instanceof Date) {
                        value = parseDate((Date) value);
                    }
                }
            }
        }

        XSSFCellStyle style = getStyleBody(color);
        style.setDataFormat(workbook.createDataFormat().getFormat("@"));
        cell.setCellStyle(style);

        if (value != null) {
            String stringValue = value.toString();
            cell.setCellValue(new XSSFRichTextString(stringValue));
            int dataWidth = calculateWidth(stringValue);
            dataWidths.merge(tpamscolumn.getCol(), dataWidth, Math::max);
        }
    }

    // 计算宽度
    private int calculateWidth(String content) {
        if (content == null || content.isEmpty()) return DEFAULT_WIDTH;
        int width = 0;
        for (char c : content.toCharArray()) {
            width += (c >= 0x4E00 && c <= 0x9FFF) ? 2 : 1;
        }
        return Math.max(width + 2, DEFAULT_WIDTH);
    }

    // 时间格式化
    private String parseDate(Date date) {
        try {
            return sdf.format(date);
        } catch (Exception e) {
            return "";
        }
    }

    // 创建表头
    public XSSFSheet createHead(XSSFSheet sheet, int r, int c) {
        for (int i = 0; i < r; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < c; j++) {
                row.createCell(j);
            }
        }
        return sheet;
    }

    public void createHead(List<Column> cellList, XSSFSheet sheet, int rowIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        for (Column tpamscolumn : cellList) {
            int r = tpamscolumn.getRow();
            int rLen = tpamscolumn.getRLen();
            int c = tpamscolumn.getCol();
            int cLen = tpamscolumn.getCLen();
            int endR = r + rLen - (rLen > 0 ? 1 : 0);
            int endC = c + cLen - (cLen > 0 ? 1 : 0);

            String content = tpamscolumn.getContent() != null ? tpamscolumn.getContent() : "";
            XSSFCell cell = row.getCell(c);
            cell.setCellStyle(styleHead);
            cell.setCellValue(new XSSFRichTextString(content));

            int contentWidth = calculateWidth(content);
            if (cLen > 0) {
                int widthPerCol = Math.max(contentWidth / cLen, DEFAULT_WIDTH);
                for (int col = c; col <= endC; col++) {
                    headerWidths.merge(col, widthPerCol, Math::max);
                }
            } else {
                headerWidths.merge(c, contentWidth, Math::max);
            }

            if (r != endR || c != endC) {
                CellRangeAddress cra = new CellRangeAddress(r, endR, c, endC);
                sheet.addMergedRegion(cra);
                RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet);
            }

            if (tpamscolumn.isHasChildren()) {
                createHead(tpamscolumn.getCellList(), sheet, r + 1);
            }
        }
    }

    // 保存方法
    private void save(XSSFWorkbook workbook, String filePath) throws IOException {
        File file = new File(filePath);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        try (FileOutputStream fOut = new FileOutputStream(file)) {
            workbook.write(fOut);
            fOut.flush();
        }
    }

    private InputStream save(XSSFWorkbook workbook) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            workbook.write(bos);
            return new ByteArrayInputStream(bos.toByteArray());
        } finally {
            bos.close();
        }
    }

    // 获取叶子节点列
    private void getColumnList(List<Column> list, List<Column> listCol) {
        for (Column column : list) {
            if (column.getFieldName() != null) {
                listCol.add(column);
            }
            getColumnList(column.getCellList(), listCol);
        }
    }

    // Column 转换方法（未修改）
    public List<Column> columnTransformer(List<T> list) {
        List<Column> lc = new ArrayList<>();
        if (list.get(0) instanceof Map) {
            final int[] i = {1};
            for (Map<String, String> m : (List<Map<String, String>>) list) {
                m.forEach((k, val) -> {
                    Column tpamscolumn = new Column();
                    tpamscolumn.setId(String.valueOf(i[0]));
                    tpamscolumn.setPid("0");
                    tpamscolumn.setContent(k);
                    tpamscolumn.setFieldName(val);
                    lc.add(tpamscolumn);
                    i[0]++;
                });
            }
        } else {
            int i = 1;
            for (String s : (List<String>) list) {
                Column tpamscolumn = new Column();
                tpamscolumn.setId(String.valueOf(i));
                tpamscolumn.setPid("0");
                tpamscolumn.setContent(s);
                tpamscolumn.setFieldName(null);
                lc.add(tpamscolumn);
                i++;
            }
        }
        setParm(lc, "0");
        List<Column> s = TreeTool.buildByRecursive(lc, "0");
        setColNum(lc, s, s);
        return s;
    }

    public List<Column> columnTransformer(List<T> list, String id, String pid, String content, String fieldName, String rootid) throws Exception {
        List<Column> lc = new ArrayList<>();
        if (list.get(0) instanceof Map) {
            for (Map m : (List<Map>) list) {
                Column tpamscolumn = new Column();
                m.forEach((k, val) -> {
                    if (id.equals(k)) tpamscolumn.setId(String.valueOf(val));
                    if (pid.equals(k)) tpamscolumn.setPid((String) val);
                    if (content.equals(k)) tpamscolumn.setContent((String) val);
                    if (fieldName != null && fieldName.equals(k)) tpamscolumn.setFieldName((String) val);
                });
                lc.add(tpamscolumn);
            }
        } else {
            for (T t : list) {
                Column tpamscolumn = new Column();
                Class cls = t.getClass();
                Field[] fs = cls.getDeclaredFields();
                for (Field f : fs) {
                    f.setAccessible(true);
                    if (id.equals(f.getName()) && f.get(t) != null) tpamscolumn.setId(f.get(t).toString());
                    if (pid.equals(f.getName()) && f.get(t) != null) tpamscolumn.setPid(f.get(t).toString());
                    if (content.equals(f.getName()) && f.get(t) != null) tpamscolumn.setContent(f.get(t).toString());
                    if (fieldName != null && f.getName().equals(fieldName) && f.get(t) != null) tpamscolumn.setFieldName(f.get(t).toString());
                }
                lc.add(tpamscolumn);
            }
        }
        setParm(lc, rootid);
        List<Column> s = TreeTool.buildByRecursive(lc, rootid);
        setColNum(lc, s, s);
        return s;
    }

    public static void setParm(List<Column> list, String rootid) {
        int totalRow = TreeTool.getMaxStep(list);
        int totalCol = TreeTool.getDownChildren(list, rootid);
        for (Column poit : list) {
            int treeStep = TreeTool.getTreeStep(list, poit.getPid(), 0);
            poit.setTreeStep(treeStep);
            poit.setRow(treeStep);
            boolean hasCh = TreeTool.hasChild(list, poit);
            poit.setHasChildren(hasCh);
            poit.setRLen(hasCh ? 0 : totalRow - treeStep);
            poit.setTotalRow(totalRow);
            poit.setTotalCol(totalCol);
        }
    }

    public static void setColNum(List<Column> list, List<Column> treeList, List<Column> flist) {
        List<Column> new_list = new ArrayList<>();
        for (Column poit : treeList) {
            int col = TreeTool.getFCol(list, poit.getPid()).getCol();
            int brotherCol = TreeTool.getBrotherChilNum(list, poit);
            poit.setCol(col + brotherCol);
            int cLen = TreeTool.getDownChildren(list, poit.getId());
            poit.setCLen(cLen <= 1 ? 0 : cLen);
            new_list.addAll(poit.getCellList());
        }
        if (!new_list.isEmpty()) {
            setColNum(list, new_list, flist);
        }
    }

    // 以下是导入Excel相关方法，未修改，保持原样
    public static String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case NUMERIC:
                case FORMULA: {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                    } else {
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case STRING:
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                default:
                    cellvalue = "";
            }
        }
        return cellvalue;
    }

    public static Workbook getWorkbookType(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(fileType)) {
            wb = new HSSFWorkbook(inStr);
        } else if (".xlsx".equals(fileType)) {
            wb = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("导入格式错误");
        }
        return wb;
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue()).trim();
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue()).trim();
            default: return "";
        }
    }

    private boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
                return true;
            }
        }
        return false;
    }

    private String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getStringCellValue(fCell);
                }
            }
        }
        return "";
    }

    /**
     * 获取excel的值 返回的 List<List<String>>的数据结构
     *
     * @param fileUrl  文件路径
     * @param sheetNum 工作表（第几分页[1,2,3.....]）
     * @return List<List < String>>
     */
    public List<List<String>> getExcelValues(String fileUrl, int sheetNum) throws Exception {
        List<List<String>> values = new ArrayList<List<String>>();
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount = sheetNum - 1; //workbook.getNumberOfSheets();//sheet 数量,可以只读取手动指定的sheet页
        //int sheetCount1= workbook.getNumberOfSheets();
        Sheet sheet = workbook.getSheetAt(sheetCount); //读取第几个工作表sheet
        int rowNum = sheet.getLastRowNum();//有多少行
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//第i行
            if (row == null) {//过滤空行
                continue;
            }
            List<String> list = new ArrayList<>();
            int colCount = sheet.getRow(0).getLastCellNum();//用表头去算有多少列，不然从下面的行计算列的话，空的就不算了
            for (int j = 0; j < colCount; j++) {//第j列://+1是因为最后一列是空 也算进去
                Cell cell = row.getCell(j);
                String cellValue;
                boolean isMerge = false;
                if (cell != null) {
                    isMerge = isMergedRegion(sheet, i, cell.getColumnIndex());
                }
                //判断是否具有合并单元格
                if (isMerge) {
                    cellValue = getMergedRegionValue(sheet, row.getRowNum(), cell.getColumnIndex());
                } else {
                    cellValue = getStringCellValue(cell);
                }
                list.add(cellValue);
            }
            values.add(list);
        }
        return values;
    }

    /**
     * 判断整行是否为空
     *
     * @param row    excel得行对象
     * @param maxRow 有效值得最大列数
     */
    private static boolean CheckRowNull(Row row, int maxRow) {
        int num = 0;
        for (int j = 0; j < maxRow; j++) {
            Cell cell = row.getCell(j);
            if (cell == null || "".equals(cell) || cell.getCellType() == CellType.BLANK) {
                num++;
            }
        }
        if (maxRow == num) {
            return true;
        }
        return false;
    }

    /**
     * 根据sheet数获取excel的值 返回List<List<Map<String,String>>>的数据结构
     *
     * @param fileUrl  文件路径
     * @param sheetNum 工作表（第几分页[1,2,3.....]）
     * @return List<List < Map < String, String>>>
     */
    public List<List<Map<String, String>>> getExcelMapVal(String fileUrl, int sheetNum) throws Exception {
        List<List<Map<String, String>>> values = new ArrayList<List<Map<String, String>>>();
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount = sheetNum - 1; //workbook.getNumberOfSheets();//sheet 数量,可以只读取手动指定的sheet页
        //int sheetCount1= workbook.getNumberOfSheets();
        Sheet sheet = workbook.getSheetAt(sheetCount); //读取第几个工作表sheet
        int rowNum = sheet.getLastRowNum();//有多少行
        Row rowTitle = sheet.getRow(0);//第i行
        int colCount = sheet.getRow(0).getLastCellNum();//用表头去算有多少列，不然从下面的行计算列的话，空的就不算了
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//第i行
            if (row == null || CheckRowNull(row, colCount)) {//过滤空行
                continue;
            }
            List<Map<String, String>> list = new ArrayList<Map<String, String>>();
            for (int j = 0; j < colCount; j++) {//第j列://+1是因为最后一列是空 也算进去
                Map<String, String> map = new HashMap<>();
                Cell cell = row.getCell(j);
                Cell cellTitle = rowTitle.getCell(j);
                String cellValue;
                String cellKey = getStringCellValue(cellTitle);
                boolean isMerge = false;
                if (cell != null) {
                    isMerge = isMergedRegion(sheet, i, cell.getColumnIndex());
                }
                //判断是否具有合并单元格
                if (isMerge) {
                    cellValue = getMergedRegionValue(sheet, row.getRowNum(), cell.getColumnIndex());
                } else {
                    cellValue = getStringCellValue(cell);
                }
                map.put(cellKey, cellValue);
                list.add(map);
            }
            values.add(list);
        }
        return values;
    }

    /**
     * 获取当前excel的工作表sheet总数
     *
     * @param fileUrl
     * @return
     * @throws Exception
     */
    public int hasSheetCount(String fileUrl) throws Exception {
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(is);
        int sheetCount = workbook.getNumberOfSheets();
        return sheetCount;
    }
}
