package com.fast.poi.excel.service;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * excel服务接口
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.service
 * @fileNmae IExcelService
 * @date 2023-8-29
 * @copyright
 * @since 0.0.1
 */

public interface IExcelService {

    Workbook export(List dataList, Class clazz);


    void exportToSheet(List dataList, Class clazz, Workbook workbook, int sheetIndex, int startRowNum);
}
