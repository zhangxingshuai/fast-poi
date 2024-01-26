package com.fast.poi.util;

import com.fast.poi.excel.bean.BaseStyle;
import com.fast.poi.excel.bean.DropDownBox;
import com.fast.poi.excel.bean.ExportField;
import com.fast.poi.excel.bean.GlobalStyle;
import com.fast.poi.excel.constants.AlignType;
import com.fast.poi.excel.context.ExportContext;
import org.apache.poi.hpsf.Decimal;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;
import java.util.Objects;


/**
 * excel工具类
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae fast.poi.common.util
 * @fileNmae ExcelUtil
 * @date 2023-8-29
 * @copyright
 * @since 0.0.1
 */
public class ExcelUtil {


    /** 给单元格设置数据
     * @param cell 单元格对象
     * @param data 数据对象
     * @param exportField 导出字段对象
     * */
    public static void setCellValue(Cell cell, Object data, ExportField exportField) {
        Object value = getValue(data, exportField.getFieldName());
        if (value == null) {
            return;
        }
        // 下拉框设置数值
        DropDownBox dropDownBox = exportField.getDropDownBox();
        if (dropDownBox != null && dropDownBox.getActive() && dropDownBox.getValuesMap().length > 0) {
            for (int i = 0; i < dropDownBox.getValuesMap().length; i++) {
                if (String.valueOf(value).replaceAll(" ","").equals(dropDownBox.getValuesMap()[i]) && i < dropDownBox.getValues().length) {
                    cell.setCellValue(dropDownBox.getValues()[i]);
                    return;
                }
            }
        }
        if (isString(value)) {
            cell.setCellValue(String.valueOf(value));
        } else if (isNumber(value)) {
            cell.setCellValue(Double.valueOf(value+""));
        } else if (isDate(value)) {
            cell.setCellValue((Date) value);
        }
    }

    /** 合并单元格
     * @param sheet sheet页对象
     * @param startRow 起始合并行
     * @param endRow 结束合并行
     * @param startCell 起始合并列
     * @param endCell 结束合并列
     * */

    public static void mergeCell(Sheet sheet, Integer startRow, Integer endRow, Integer startCell, Integer endCell) {
        if (Objects.equals(startRow, endRow) && Objects.equals(startCell, endCell)) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(startRow,endRow,startCell,endCell);
        sheet.addMergedRegion(region);
    }

    /** 反射获取对象指定属性的属性值 */
    private static Object getValue(Object data, String fieldName) {
        Class<?> clazz = data.getClass();
        try {
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            return field.get(data);
        } catch (Exception e) {
            return null;
        }
    }


    /** 判断数据是否为日期类型 */
    public static boolean isDate(Object value) {
        return Date.class.getName().equals(value.getClass().getName());
    }

    /** 判断数据是否为字符串类型*/
    public static boolean isString(Object value){
        return String.class.getName().equals(value.getClass().getTypeName());
    }

    /** 判断数据是否为数值类型 */
    public static boolean isNumber(Object value){
        String type = value.getClass().getName();
        if (Integer.class.getName().equals(type) || int.class.getName().equals(type)
        || Double.class.getName().equals(type) || double.class.getName().equals(type)
        || Long.class.getName().equals(type) || long.class.getName().equals(type)
        || Short.class.getName().equals(type) || short.class.getName().equals(type)
        || Float.class.getName().equals(type) || float.class.getName().equals(type)
        || Decimal.class.getName().equals(type)) {
            return true;
        }
        return false;
    }


    /** 设置单元格样式 */
    public static void setStyle(Cell cell, ExportField exportField, List<BaseStyle> styleList) {
        CellStyle style = ExportContext.getWorkbook().createCellStyle();
        if (styleList == null || styleList.size() == 0) {
            return;
        }
        for (BaseStyle baseStyle : styleList) {

        }
//        // 全局样式
//        if (globalStyle != null) {
//            ExcelUtil.setBorderStyle(style, globalStyle);
//            ExcelUtil.setFontStyle(style, globalStyle);
//            ExcelUtil.setBgStyle(style, globalStyle);
//        }
        // 设置自动换行
        style.setWrapText(true);
        // 设置对齐方式
        setAlign(style,exportField.getAlign());
        cell.setCellStyle(style);
    }

    /** 设置背景颜色 */
    public static void setBgStyle(CellStyle style, GlobalStyle globalStyle) {
        if (globalStyle.getBgColor().equals(IndexedColors.WHITE)) {
            return;
        }
        // 设置背景色
        XSSFColor color = new XSSFColor();
        color.setIndexed(globalStyle.bgColor.getIndex());
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    /** 设置字体样式 */
    public static void setFontStyle(CellStyle style, GlobalStyle globalStyle) {
        Font font = ExportContext.getWorkbook().createFont();
        font.setFontHeightInPoints(globalStyle.getFontSize());
        font.setBold(globalStyle.getBold());
        font.setFontName(globalStyle.getFontName());
        style.setFont(font);
    }

    /** 设置边框样式 */
    public static void setBorderStyle(CellStyle style, GlobalStyle globalStyle) {
        if (globalStyle.getShowBorder()) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
        }
    }

    /** 设置对齐方式 */
    public static void setAlign(CellStyle style, AlignType alignType) {
        if (alignType == AlignType.CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (alignType == AlignType.RIGHT) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        }
        // 垂直默认居中对齐
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }

}
