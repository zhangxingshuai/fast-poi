package com.fast.poi.excel.service.impl;

import com.fast.poi.excel.bean.BaseStyle;
import com.fast.poi.excel.bean.GlobalStyle;
import com.fast.poi.excel.context.ExportContext;
import com.fast.poi.excel.service.IStyleParser;
import com.fast.poi.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.springframework.stereotype.Service;

/**
 * 全局样式解析器
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.service.impl
 * @fileNmae SimpleStyleParserImpl
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */

@Service
public class GlobalStyleParserImpl implements IStyleParser {

    @Override
    public boolean support(BaseStyle baseStyle) {
        return baseStyle instanceof GlobalStyle;
    }

    @Override
    public void parse(Cell cell, BaseStyle baseStyle) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = ExportContext.getWorkbook().createCellStyle();
        }
        ExcelUtil.setBorderStyle(cellStyle, (GlobalStyle) baseStyle);
        ExcelUtil.setFontStyle(cellStyle, (GlobalStyle) baseStyle);
        ExcelUtil.setBgStyle(cellStyle, (GlobalStyle) baseStyle);
        cell.setCellStyle(cellStyle);
    }
}
