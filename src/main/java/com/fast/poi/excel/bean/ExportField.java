package com.fast.poi.excel.bean;

import com.fast.poi.excel.constants.AlignType;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae ExcelField
 * @date 2023-8-30
 * @copyright
 * @since 0.0.1
 */
@Data
public class ExportField {

    /** 列索引 */
    private Integer columnIndex;

    /** 字段名称 */
    private String fieldName;

    /** 列宽 */
    private Integer width;

    /** 对齐方式 */
    private AlignType align;

    /** 是否需要行合并 */
    private boolean mergeRow;

    /** 指定合并依赖字段，与指定字段行合并保持一致  */
    private String mergeDepend;

    /** 行合并长度 */
    private Integer mergeRowLen;

    /** 单元格下拉框 */
    private DropDownBox dropDownBox;

    /** 日期格式 */
    private String format;

    /** 单元格背景色 */
    private IndexedColors bgColor;

}
