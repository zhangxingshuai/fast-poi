package com.fast.poi.excel.constants;

import lombok.Getter;

/**
 * 条件枚举
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.constants
 * @fileNmae Condition
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */
@Getter
public enum Condition {

    EQ("等于", "="),
    GT("大于", ">"),
    GE("大于等于", ">="),
    LT("小于", "<"),
    LE("小于等于", "<="),
    NE("不等于", "!="),
    AND("并且", "and"),
    OR("或者", "or");

    private String description;
    private String value;

    Condition(String description, String value) {
        this.description = description;
        this.value = value;
    }
}
