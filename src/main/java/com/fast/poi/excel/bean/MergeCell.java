package com.fast.poi.excel.bean;

import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 合并单元格
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae MergeCell
 * @date 2023-8-30
 * @copyright
 * @since 0.0.1
 */
@Data
@NoArgsConstructor
public class MergeCell implements Cloneable {

    private Integer startRow;

    private Integer endRow;

    private Integer startCell;

    private Integer endCell;

    public MergeCell(Integer startRow, Integer endRow, Integer startCell, Integer endCell) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCell = startCell;
        this.endCell = endCell;
    }
}
