package com.fast.poi.excel.constants;

/**
 * excel 对齐方式枚举
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.common.constants
 * @fileNmae AlignType
 * @date 2023-8-29
 * @copyright
 * @since 0.0.1
 */
public enum AlignType {

    LEFT("左对齐"),
    CENTER("居中对齐"),
    RIGHT("右对齐");

    private String description;

    AlignType(String description) {
        this.description = description;
    }
}
