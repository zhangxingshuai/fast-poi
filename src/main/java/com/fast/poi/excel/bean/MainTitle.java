package com.fast.poi.excel.bean;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 主标题
 *
 * @author 张兴帅
 * @projectName jijian
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae MainTitle
 * @date 2024/1/3
 * @copyright
 * @since 0.0.1
 */
@Data
@AllArgsConstructor
public class MainTitle {

    /** 标题 */
    private String title;

    /** 起始合并列索引 */
    private Integer startCellIndex;

    /** 结束合并列索引 */
    private Integer endCellIndex;
}
