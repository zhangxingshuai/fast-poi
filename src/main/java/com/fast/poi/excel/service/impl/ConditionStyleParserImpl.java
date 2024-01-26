package com.fast.poi.excel.service.impl;

import com.fast.poi.excel.bean.BaseStyle;
import com.fast.poi.excel.bean.ConditionStyle;
import com.fast.poi.excel.constants.Condition;
import com.fast.poi.excel.context.ExportContext;
import com.fast.poi.excel.service.IStyleParser;
import com.fast.poi.util.ExcelUtil;
import com.fast.poi.util.TimeUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;

/**
 * 条件样式解析器
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.service.impl
 * @fileNmae ConditionStyleParserImpl
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */

@Service
@Slf4j
public class ConditionStyleParserImpl implements IStyleParser {
    @Override
    public boolean support(BaseStyle baseStyle) {
        return baseStyle instanceof ConditionStyle;
    }

    @Override
    public void parse(Cell cell, BaseStyle baseStyle) {
        // 如果当前行数据对象为空，不设置样式
        if (ExportContext.getCurrentData() == null) {
            return;
        }
        ConditionStyle style = (ConditionStyle) baseStyle;
        // 如果当前单元格不是条件样式的列
        if (cell.getColumnIndex() != style.getColumnIndex()) {
            return;
        }
        // 如果条件表达式校验通过，进行样式设置
        if (checkCondition(style.getCondition())) {
            setStyle(cell, style);
        }
    }

    /** 设置单元格样式
     * @param cell 单元格对象
     * @param style 样式对象
     * */
    private void setStyle(Cell cell, ConditionStyle style) {
        CellStyle cellStyle = cell.getCellStyle();
        // 设置背景色
        cellStyle.setFillForegroundColor(style.bgColor.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置字体大小
        Font font = ExportContext.getWorkbook().createFont();
        font.setFontHeightInPoints(style.getFontSize());
        // 字体是否加粗
        font.setBold(style.getBold());
        // 字体颜色
        font.setColor(style.getFontColor().getIndex());
        cellStyle.setFont(font);

    }

    /** 校验条件表达式
     * @param condition 条件表达式
     * @return true 校验成功  false 校验失败
     * */
    private boolean checkCondition(String condition) {
        if (condition == null || condition.replaceAll(" ","").equals("")) {
            return false;
        }
        // 判断是否存在and、or条件
        if (condition.indexOf(Condition.AND.getValue()) != -1 || condition.indexOf(Condition.OR.getValue()) != -1) {
            return checkMultiCondition(condition);
        } else { // 单条件判断
            return checkSingleCondition(condition);
        }
    }

    // TODO 目前多条件检查只支持同类型条件，暂不支持混合条件及其他复杂条件，后续需要完善
    /** 多条件检查 */
    private boolean checkMultiCondition(String condition) {
        String[] arrAnd = condition.split(Condition.AND.getValue());
        String[] arrOr = condition.split(Condition.OR.getValue());
        // 且判断
        if (arrAnd.length > 1) {
            for (String conditionPart : arrAnd) {
                if (!checkSingleCondition(conditionPart)) { // 有假则假
                    return false;
                }
            }
            return true;
        }
        // 或判断
        if (arrOr.length > 1) {
            for (String conditionPart : arrOr) {
                if (checkSingleCondition(conditionPart)) { // 有真则真
                    return true;
                }
            }
            return false;
        }
        return false;
    }

    private boolean checkSingleCondition(String condition) {
        for (Condition conditionType : Condition.values()) {
            String[] arr = condition.split(conditionType.getValue());
            if (arr.length > 1) {  // 命中条件
                // 条件
                String fieldName = arr[0].replaceAll(" ","");
                String value = arr[1].replaceAll(" ","");
                return doCheck(fieldName, conditionType, value);
            }
        }
        return false;
    }

    private boolean doCheck(String fieldName, Condition conditionType, String value) {
        // 获取真实属性值
        Object currentData = ExportContext.getCurrentData();
        try {
            Field field = currentData.getClass().getDeclaredField(fieldName);
            field.setAccessible(true);
            Object realValue = field.get(currentData);
            if (realValue == null) { // 空值校验
                return checkNull(realValue,value,conditionType);
            }
            // 非空值校验
            if (ExcelUtil.isString(realValue)) { // 字符串类型比较
                return checkString(realValue,value,conditionType);
            } else if (ExcelUtil.isNumber(realValue)) { // 数值类型比较
                return checkNumber(realValue,value,conditionType);
            } else if (ExcelUtil.isDate(realValue)) { // 日期类型比较
                return checkDate(realValue,value,conditionType);
            }
        } catch (Exception e) {
            log.error("excel条件样式解析异常-字段名称={}", fieldName);
            return false;
        }
        return false;

    }

    /** 空值校验 */
    private boolean checkNull(Object realValue, String value, Condition conditionType) {
        if (conditionType == Condition.EQ && "null".equals(value.toLowerCase()) || "''".equals(value)) {
            return true;
        }
        return false;
    }


    /** 字符串类型校验 */
    private boolean checkString(Object realValue, String value, Condition conditionType) {
        if (conditionType == Condition.EQ) {
            if (realValue.equals(value.replaceAll("'","")) || realValue.equals("") && value.equals("''")) {
                return true;
            }
        }
        if (conditionType == Condition.NE) {
            return !realValue.equals(value);
        }
        if (conditionType == Condition.GT) {
            return realValue.hashCode() > value.hashCode();
        }
        if (conditionType == Condition.GE) {
            return realValue.hashCode() >= value.hashCode();
        }
        if (conditionType == Condition.LT) {
            return realValue.hashCode() < value.hashCode();
        }
        if (conditionType == Condition.LE) {
            return realValue.hashCode() <= value.hashCode();
        }
        return false;
    }

    /** 数值类型校验 */
    private boolean checkNumber(Object realValue, String value, Condition conditionType) {

        if (conditionType == Condition.EQ) {
            return Double.valueOf(realValue+"").doubleValue() == Double.valueOf(value).doubleValue();
        }
        if (conditionType == Condition.NE) {
            return Double.valueOf(realValue+"").doubleValue() != Double.valueOf(value).doubleValue();
        }
        if (conditionType == Condition.GT) {
            return Double.valueOf(realValue+"").doubleValue() > Double.valueOf(value).doubleValue();
        }
        if (conditionType == Condition.GE) {
            return Double.valueOf(realValue+"").doubleValue() >= Double.valueOf(value).doubleValue();
        }
        if (conditionType == Condition.LT) {
            return Double.valueOf(realValue+"").doubleValue() < Double.valueOf(value).doubleValue();
        }
        if (conditionType == Condition.LE) {
            return Double.valueOf(realValue+"").doubleValue() <= Double.valueOf(value).doubleValue();
        }
        return false;
    }

    /** 日期类型校验 */
    private boolean checkDate(Object realValue, String value, Condition conditionType) {
        long stamp1 = TimeUtil.getTimestamp(realValue.toString());
        long stamp2 = TimeUtil.getTimestamp(value.toString());
        if (conditionType == Condition.EQ) {
            return stamp1 == stamp2;
        }
        if (conditionType == Condition.NE) {
            return stamp1 != stamp2;
        }
        if (conditionType == Condition.GT) {
            return stamp1 > stamp2;
        }
        if (conditionType == Condition.GE) {
            return stamp1 >= stamp2;
        }
        if (conditionType == Condition.LT) {
            return stamp1 < stamp2;
        }
        if (conditionType == Condition.LE) {
            return stamp1 <= stamp2;
        }
        return false;
    }


}
