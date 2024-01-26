package com.fast.poi.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 条件样式注解
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.annotation
 * @fileNmae ConditionStyle
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelConditionStyle {

    /** 条件表达式 */
    String condition() default "";

    /** 背景颜色，默认白色 */
    IndexedColors bgColor() default IndexedColors.WHITE;

    /** 字体大小，默认11 */
    short fontSize() default 11;

    /** 字体是否加粗，默认否 */
    boolean bold() default false;

    /** 字体颜色，默认黑色 */
    IndexedColors fontColor() default IndexedColors.BLACK;
}
