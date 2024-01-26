package com.fast.poi.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel导出全局样式注解
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.annotation
 * @fileNmae ExcelGlobalStyle
 * @date 2023-8-30
 * @copyright
 * @since 0.0.1
 */

@Retention(value = RetentionPolicy.RUNTIME)
@Target(value = ElementType.TYPE)
public @interface ExcelGlobalStyle {

    /** 是否展示边框，默认不展示*/
    boolean showBorder() default false;

    /** 字体大小，默认11 */
    short fontSize() default 11;

    /** 字体是否加粗，默认否 */
    boolean bold() default false;

    /** 字体样式，默认宋体 */
    String fontName() default "宋体";

    /** 背景色，默认白色 */
    IndexedColors bgColor() default IndexedColors.WHITE;

    /** 行高，默认20 */
    float rowHeight() default 20;
}
