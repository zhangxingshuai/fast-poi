package com.fast.poi.excel.service.impl;

import com.fast.poi.excel.annotation.ExcelDropDownBox;
import com.fast.poi.excel.annotation.ExcelConditionStyle;
import com.fast.poi.excel.annotation.ExcelField;
import com.fast.poi.excel.annotation.ExcelGlobalStyle;
import com.fast.poi.excel.context.ExportContext;
import com.fast.poi.excel.service.IExcelService;
import com.fast.poi.excel.service.IStyleParser;
import com.fast.poi.util.ExcelUtil;
import com.fast.poi.excel.bean.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * excel导出实现类
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.service.impl
 * @fileNmae ExcelServiceImpl
 * @date 2023-8-29
 * @copyright
 * @since 0.0.1
 */
@Service
public class ExcelServiceImpl implements IExcelService {

    @Resource
    private List<IStyleParser> styleParserList;

    /** 导出excel数据：返回excel工作簿对象
     * @param dataList 数据集合
     * @param clazz 需要导出的类对象
     * */
    @Override
    public  Workbook export(List dataList,Class clazz){
        // 创建excel工作簿对象
        Workbook workbook = createWorkbook();
        // 解析样式
        resolveStyle(clazz);
        // 创建sheet页
        Sheet sheet = workbook.createSheet();
        // 创建标题
        createTitle(sheet.createRow(0), clazz);
        // 创建数据行
        createDataRow(sheet,dataList);
        // 合并单元格
        mergeCell(sheet,dataList);
        // 设置下拉选
        setDropDownBox(sheet,dataList.size());
        // 设置单元格整体列宽
        setColumnWidth(sheet);
        // 清除excel导出上下文数据
        ExportContext.clear();
        return workbook;
    }

    /** 将数据导出到指定sheet页的指定行 */
    @Override
    public void exportToSheet(List dataList, Class clazz, Workbook workbook, int sheetIndex, int startRowNum) {
        ExportContext.setWorkbook(workbook);
        ExportContext.setStartRowNum(startRowNum);
        // 解析样式
        resolveStyle(clazz);
        // 创建sheet页
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        // 创建标题
        createTitle(sheet.createRow(startRowNum), clazz);
        // 创建数据行
        createDataRow(sheet,dataList);
        // 合并单元格
        mergeCell(sheet,dataList);
        // 设置下拉选
        setDropDownBox(sheet,dataList.size());
        // 设置单元格整体列宽
        setColumnWidth(sheet);
        // 清除excel导出上下文数据
        ExportContext.clear();
    }

    /** 设置单元格下拉框 */
    private void setDropDownBox(Sheet sheet, int dataSize) {
        List<ExportField> exportFields = ExportContext.getExportFields();
        for (ExportField exportField : exportFields) {
            DropDownBox dropDownBox = exportField.getDropDownBox();
            if (dropDownBox != null && dropDownBox.getActive()) {
                // 创建下拉选
                XSSFDataValidationHelper helper = new XSSFDataValidationHelper((XSSFSheet) sheet);
                XSSFDataValidationConstraint constraint = new XSSFDataValidationConstraint(dropDownBox.getValues());
                CellRangeAddressList addressList = new CellRangeAddressList(ExportContext.getStartRowNum(), ExportContext.getStartRowNum()+dataSize, exportField.getColumnIndex(), exportField.getColumnIndex());
                DataValidation validation = helper.createValidation(constraint, addressList);
                // 设置单元格只能是下拉选中的数据，否则提示错误
                validation.setSuppressDropDownArrow(true);
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
        }
    }

    /** 创建工作簿对象，并保存到上下文 */
    private Workbook createWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 设置工作簿对象到上下文
        ExportContext.setWorkbook(workbook);
        return workbook;
    }

    /** 解析样式 */
    private void resolveStyle(Class clazz) {
        // 解析全局样式
        resolveGlobalStyle(clazz);
        // 解析条件样式
        resolveConditionStyle(clazz);
    }

    /** 解析条件样式 */
    private void resolveConditionStyle(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (int i=0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.isAnnotationPresent(ExcelConditionStyle.class)) {
                ExcelConditionStyle annotation = (ExcelConditionStyle) field.getAnnotation(ExcelConditionStyle.class);
                ConditionStyle conditionStyle = new ConditionStyle();
                conditionStyle.setColumnIndex(i);
                conditionStyle.setCondition(annotation.condition());
                conditionStyle.setBgColor(annotation.bgColor());
                conditionStyle.setFontColor(annotation.fontColor());
                conditionStyle.setFontSize(annotation.fontSize());
                conditionStyle.setBold(annotation.bold());
                ExportContext.addBaseStyle(conditionStyle);
            }
        }
    }

    /** 解析excel全局样式 */
    private void resolveGlobalStyle(Class clazz) {
        if (clazz.isAnnotationPresent(ExcelGlobalStyle.class)) {
            ExcelGlobalStyle annotation = (ExcelGlobalStyle) clazz.getAnnotation(ExcelGlobalStyle.class);
            GlobalStyle globalStyle = new GlobalStyle();
            globalStyle.setShowBorder(annotation.showBorder());
            globalStyle.setFontSize(annotation.fontSize());
            globalStyle.setFontName(annotation.fontName());
            globalStyle.setBold(annotation.bold());
            globalStyle.setBgColor(annotation.bgColor());
            globalStyle.setRowHieght(annotation.rowHeight());
            ExportContext.addBaseStyle(globalStyle);
        }
    }

    /** 合并单元格（目前只支持行合并，后续可根据需求扩展列合并） */
    private void mergeCell(Sheet sheet, List dataList) {
        List<MergeCell> mergeCellList = new ArrayList<>();
        // 筛选出所有需要合并的字段
        List<ExportField> fieldList = ExportContext.getExportFields().stream().filter(ExportField::isMergeRow).collect(Collectors.toList());
        // 存放每列需要进行合并的片段
        Map<String,List<MergeCell>> mergeMap = new HashMap<>();
        // 计算出所有合并对象
        for (ExportField exportField : fieldList) {
            if (StringUtils.isNotEmpty(exportField.getMergeDepend())) { // 如果依赖合并字段不为空，则与依赖字段合并行保持一致
                List<MergeCell> dependMergeList = mergeMap.get(exportField.getMergeDepend());
                if (dependMergeList != null) {
                    List<MergeCell> list = new ArrayList<>(dependMergeList.size());
                    for (MergeCell mergeCell : dependMergeList) { // 修改合并列为当前字段所在列
                        list.add(new MergeCell(mergeCell.getStartRow(),mergeCell.getEndRow(),exportField.getColumnIndex(),exportField.getColumnIndex()));
                    }
                    mergeMap.put(exportField.getFieldName(),list);
                    mergeCellList.addAll(list);
                }
            } else {
                List mergeList = calculMergeCell(dataList, exportField);
                mergeMap.put(exportField.getFieldName(), mergeList);
                mergeCellList.addAll(mergeList);
            }
        }
        // 循环合并
        for (MergeCell mergeCell : mergeCellList) {
//            if (mergeCell.getStartRow() == mergeCell.getEndRow()) {
//                continue;
//            }
            ExcelUtil.mergeCell(sheet, mergeCell.getStartRow(),mergeCell.getEndRow(),mergeCell.getStartCell(),mergeCell.getEndCell());
        }
    }

    /** 计算合并单元格（计算逻辑：初始起始行、结束行，结束数据与起始行数据比较，如果一致则结束行向下移动，继续比较，直到结束行数据与起始行数据不一致，则认为需要进行合并） */
    private List<MergeCell> calculMergeCell(List<Object> dataList, ExportField exportField) {
        List<MergeCell> mergeCellList = new ArrayList<>();
        // 起始合并行
        int start = 0;
        for (int end = 0; end < dataList.size(); end++) {
            // 如果当前行满足合并需求
            if (needMerge(dataList.get(start),dataList.get(end),exportField.getFieldName())) {
                if (start == end-1) {
                    start = end;
                    continue;
                }
                mergeCellList.add(buildMergeCell(start+ExportContext.getStartRowNum(),end-1+ExportContext.getStartRowNum(),exportField.getColumnIndex()));
                start = end;
            } else if (end == dataList.size()-1 && start != end-1) { // 判断最后一行是否可与start行合并
                mergeCellList.add(buildMergeCell(start+ExportContext.getStartRowNum(),end+ExportContext.getStartRowNum(),exportField.getColumnIndex()));
            }
        }

        return mergeCellList;
    }

    /** 构建单元格合并对象
     * @param start 起始合并行
     * @param end 结束合并行
     * @param columnIndex 列索引
     * */
    private MergeCell buildMergeCell(int start, int end, Integer columnIndex) {
        MergeCell mergeCell = new MergeCell();
        mergeCell.setStartRow(start);
        mergeCell.setEndRow(end);
        mergeCell.setStartCell(columnIndex);
        mergeCell.setEndCell(columnIndex);
        return mergeCell;
    }

    /** 判断是否满足合并条件：起始行数据与当前行数据不同时，满足合并条件
     * @param startObj 起始行数据对象
     * @param endObj 当前行数据对象
     * @param fieldName 字段名称
     * */
    private boolean needMerge(Object startObj, Object endObj, String fieldName) {
        Class<?> clazz = startObj.getClass();
        try {
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            Object startVal = field.get(startObj);
            Object endVal = field.get(endObj);
            if (startVal != null && endVal != null) {
                return !startVal.equals(endVal);
            } else if ((startVal == null && endVal != null) || (startVal != null && endVal == null)) {
                return true;
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return false;
    }

    /** 创建数据行
     * @param sheet sheet页对象
     * @param dataList 数据列表
     * */
    private void createDataRow(Sheet sheet, List<Object> dataList) {
        for (int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(i+ExportContext.getStartRowNum());
            row.setHeightInPoints(40);
            ExportContext.setCurrentData(dataList.get(i));
            // 创建数据列
            List<ExportField> fieldList = ExportContext.getExportFields();
            for (ExportField exportField : fieldList) {
                // 创建单元格
                Cell cell = row.createCell(exportField.getColumnIndex());
                // 设置单元格数据
                ExcelUtil.setCellValue(cell,dataList.get(i),exportField);
                // 设置单元格样式
                setStyle(cell, exportField,ExportContext.getBaseStyle());
            }
        }

    }

    /** 设置单元格列宽 */
    private void setColumnWidth(Sheet sheet) {
        List<ExportField> exportFieldList = ExportContext.getExportFields();
        for (ExportField exportField : exportFieldList) {
            Integer columnIndex = exportField.getColumnIndex();
            sheet.setColumnWidth(columnIndex, exportField.getWidth());
        }
    }

    /** 创建标题并构建导出字段 */
    private void createTitle(Row row, Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        // 主标题
        List<MainTitle> mainTitles = new ArrayList<>();
        // 副标题
        List<String> titles = new ArrayList<>();
        int index = 0;
        for (int i=0; i<fields.length; i++) {
            Field field = fields[i];
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField annotation = field.getAnnotation(ExcelField.class);
                if (StringUtils.isNotEmpty(annotation.mainTitle())) { // 主标题不为空
                   if (mainTitles.size() == 0) {
                       mainTitles.add(new MainTitle(annotation.mainTitle(), i, i));
                   } else {
                       MainTitle mainTitle = mainTitles.get(mainTitles.size()-1);
                       if (mainTitle.getTitle().equals(annotation.mainTitle())) {
                           mainTitle.setEndCellIndex(i);
                       } else {
                           mainTitles.add(new MainTitle(annotation.mainTitle(), i, i));
                       }
                   }
                }
                titles.add(annotation.title());
                // 构建导出字段
                ExportField exportField = buildExportField(field, annotation, index);
                ExportContext.addExportField(exportField);
                index++;
            }
        }
        // 设置主标题
        if (mainTitles.size() > 0) {
            // 设置数据起始行
            ExportContext.setStartRowNum(2);
            Row mainTitleRow = ExportContext.getWorkbook().getSheetAt(0).createRow(0);
            for (MainTitle mainTitle : mainTitles) {
                for (int i = mainTitle.getStartCellIndex(); i <= mainTitle.getEndCellIndex(); i++) {
                    Cell cell = mainTitleRow.createCell(i);
                    cell.setCellValue(mainTitle.getTitle());
                    // 设置标题样式
                    setTitleStyle(cell);
                }
                ExportContext.getWorkbook().getSheetAt(0).addMergedRegion(new CellRangeAddress(0,0,mainTitle.getStartCellIndex(),mainTitle.getEndCellIndex()));
            }
        }
        // 设置标题
        Row titleRow = ExportContext.getWorkbook().getSheetAt(0).createRow(mainTitles.size()==0?0:1);
        for (int i = 0; i < titles.size(); i++) {
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(titles.get(i));
            // 设置标题样式
            setTitleStyle(cell);
        }

    }

    /** 设置样式 */
    private void setStyle(Cell cell, ExportField exportField, List<BaseStyle> styleList) {
        // 字段样式设置
        CellStyle style = ExportContext.getWorkbook().createCellStyle();
        // 设置自动换行
        style.setWrapText(true);
        // 设置对齐方式
        ExcelUtil.setAlign(style,exportField.getAlign());
        // todo 设置单元格背景颜色后，会导致excel默认边框消失，暂时先不设置背景色
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(exportField.getBgColor().getIndex());
        cell.setCellStyle(style);
        // 设置日期格式
        if (StringUtils.isNotEmpty(exportField.getFormat())) {
            short format = ExportContext.getWorkbook().createDataFormat().getFormat(exportField.getFormat());
            style.setDataFormat(format);
        }
        if (styleList != null && styleList.size() > 0) {
            // 设置全局样式、条件样式
            for (BaseStyle baseStyle : styleList) {
                for (IStyleParser parser : styleParserList) {
                    if (parser.support(baseStyle)) {
                        parser.parse(cell, baseStyle);
                    }
                }
            }
        }

    }

    /** 构建导出字段 */
    private ExportField buildExportField(Field field, ExcelField annotation, int index) {
        ExportField exportField = new ExportField();
        exportField.setColumnIndex(index);
        exportField.setFieldName(field.getName());
        exportField.setWidth(annotation.width());
        exportField.setAlign(annotation.align());
        exportField.setMergeRow(annotation.mergeRow());
        exportField.setMergeDepend(annotation.mergeDepend());
        exportField.setFormat(annotation.format());
        exportField.setBgColor(annotation.bgColor());
        if (field.isAnnotationPresent(ExcelDropDownBox.class)) {
            ExcelDropDownBox dropDownAnnotaion = field.getAnnotation(ExcelDropDownBox.class);
            DropDownBox dropDownBox = new DropDownBox();
            dropDownBox.setActive(dropDownAnnotaion.active());
            dropDownBox.setValues(dropDownAnnotaion.values());
            dropDownBox.setValuesMap(dropDownAnnotaion.valuesMap());
            exportField.setDropDownBox(dropDownBox);
        }
        return exportField;
    }

    /** 设置头部样式 */
    private void setTitleStyle(Cell cell) {
        CellStyle style = ExportContext.getWorkbook().createCellStyle();
        // 标题居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 字体加粗
        Font font = ExportContext.getWorkbook().createFont();
        font.setBold(true);
        style.setFont(font);
        // 自动换行
        style.setWrapText(true);
        // 边框
        List<BaseStyle> styles = ExportContext.getBaseStyle();
        for (BaseStyle baseStyle : styles) {
            if (baseStyle instanceof GlobalStyle) {
                ExcelUtil.setBorderStyle(style,(GlobalStyle)baseStyle);
            }
        }
        cell.setCellStyle(style);

    }


}
