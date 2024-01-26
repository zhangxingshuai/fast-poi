package com.fast.poi.excel.service;

import com.fast.poi.excel.bean.BaseStyle;
import org.apache.poi.ss.usermodel.Cell;

/**
 * excel样式解析器接口
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.service
 * @fileNmae IStyleParser
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */
public interface IStyleParser {

    boolean support(BaseStyle baseStyle);

    void parse(Cell cell, BaseStyle style);
}
