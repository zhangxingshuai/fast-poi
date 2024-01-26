package com.fast.poi.word.constants;

import lombok.Getter;

/**
 * word 文档占位符
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.constants
 * @fileNmae Placeholder
 * @date 2023-9-5
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Getter
public enum Placeholder {

    /** 开始标识 */
    START("\\$\\{"),
    /** 结束标识 */
    END("\\}"),

    /** 循环体开始标识*/
    LOOP_START("<LOOP:"),

    /** 循环体结束标识*/
    LOOP_END("</LOOP>"),

    /** 循环体内部占位符开始标识 */
    LOOP_TEXT_START("#\\{"),

    /** 循环体内部占位符结束标识 */
    LOOP_TEXT_END("\\}");

    private String value;

    Placeholder(String value) {
        this.value = value;
    }
}
