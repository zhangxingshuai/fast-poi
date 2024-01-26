package com.fast.poi.excel.bean;

import lombok.Data;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * excel上下文对象
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae ExcelContext
 * @date 2023-9-1
 * @copyright
 * @since 0.0.1
 */
@Data
public class ExcelContext {

    /** 工作簿对象 */
    private Workbook workbook;

    /** 导出字段列表 */
    private List<ExportField> exportFieldList;

    /** 样式列表 */
    private List<BaseStyle> styleList;

    /** 当前行数据对象 */
    private Object currentRowData;

    /** 起始数据行索引 */
    private Integer startRowNum = 1;
}
