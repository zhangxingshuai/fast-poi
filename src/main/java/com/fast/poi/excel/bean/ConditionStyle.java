package com.fast.poi.excel.bean;

import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 条件样式
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae ConditionStyle
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */

@Data
public class ConditionStyle extends BaseStyle {

    /** 列索引 */
    private Integer columnIndex;

    /** 条件表达式 */
    private String condition;

    /** 背景颜色 */
    public IndexedColors bgColor;

    /** 字体颜色 */
    public IndexedColors fontColor;

    /** 字体大小 */
    public short fontSize;

    /** 是否加粗 */
    public Boolean bold;

}
