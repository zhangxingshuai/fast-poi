package com.fast.poi.excel.context;

import com.fast.poi.excel.bean.BaseStyle;
import com.fast.poi.excel.bean.ExcelContext;
import com.fast.poi.excel.bean.ExportField;
import com.fast.poi.excel.bean.GlobalStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * excel导出上下文
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.context
 * @fileNmae ExcelExportContext
 * @date 2023-8-30
 * @copyright
 * @since 0.0.1
 */
public class ExportContext {

    /** Excel上下文对象 */
    private static final ThreadLocal<ExcelContext> EXCEL_CONTEXT = new ThreadLocal<>();

    public static ExcelContext getExcelContext() {
        return EXCEL_CONTEXT.get();
    }

    public static void setExcelContext(ExcelContext excelContext) {
        EXCEL_CONTEXT.set(excelContext);
    }

    private static ExcelContext getExcelContextForSet() {
        ExcelContext excelContext = getExcelContext();
        excelContext = excelContext == null ? new ExcelContext() : excelContext;
        return excelContext;
    }

    /** 移除excel上下文对象 */
    public static void removeExcelContext() {
        EXCEL_CONTEXT.remove();
    }

    /** 添加导出字段列表 */
    public static void addExportField(ExportField exportField) {
        ExcelContext excelContext = getExcelContextForSet();
        List<ExportField> fieldList = excelContext.getExportFieldList()==null?new ArrayList<>():excelContext.getExportFieldList();
        fieldList.add(exportField);
        excelContext.setExportFieldList(fieldList);
        setExcelContext(excelContext);
    }

    /** 获取导出字段列表 */
    public static List<ExportField> getExportFields(){
        ExcelContext excelContext = getExcelContext();
        return excelContext == null ? null : excelContext.getExportFieldList();
    }

    /** 添加样式 */
    public static void addBaseStyle(BaseStyle baseStyle) {
        ExcelContext excelContext = getExcelContextForSet();
        List<BaseStyle> styleList = excelContext.getStyleList()==null?new ArrayList<>():excelContext.getStyleList();
        styleList.add(baseStyle);
        excelContext.setStyleList(styleList);
        setExcelContext(excelContext);
    }

    public static List<BaseStyle> getBaseStyle(){
        ExcelContext excelContext = getExcelContext();
        return  excelContext == null ? null : excelContext.getStyleList();
    }

    /** 获取全局样式 */
    public static GlobalStyle getGlobalStyle() {
        Optional<BaseStyle> optional = getBaseStyle().stream().filter(s -> s instanceof GlobalStyle).findFirst();
        if (optional.isPresent()) {
            return (GlobalStyle) optional.get();
        }
        return null;
    }

    /** 设置工作簿对象 */
    public static void setWorkbook(Workbook workbook) {
        ExcelContext excelContext = getExcelContextForSet();
        excelContext.setWorkbook(workbook);
        setExcelContext(excelContext);
    }

    /** 获取工作簿对象 */
    public static Workbook getWorkbook(){
        ExcelContext excelContext = getExcelContext();
        return excelContext == null ? null : excelContext.getWorkbook();
    }


    /** 设置当前行数据对象 */
    public static void setCurrentData(Object data) {
        ExcelContext excelContext = getExcelContextForSet();
        excelContext.setCurrentRowData(data);
        setExcelContext(excelContext);
    }

    /** 获取当前行数据对象 */
    public static Object getCurrentData() {
        ExcelContext excelContext = getExcelContext();
        return excelContext == null ? null : excelContext.getCurrentRowData();
    }

    /** 设置当前行数据对象 */
    public static void setStartRowNum(Integer startRowNum) {
        ExcelContext excelContext = getExcelContextForSet();
        excelContext.setStartRowNum(startRowNum);
        setExcelContext(excelContext);
    }

    /** 获取当前行数据对象，默认第1行 */
    public static Integer getStartRowNum() {
        ExcelContext excelContext = getExcelContext();
        return excelContext == null ? 1 : excelContext.getStartRowNum();
    }

    /** 清除excel导出上下文对象 */
    public static void clear() {
        removeExcelContext();
    }

}
