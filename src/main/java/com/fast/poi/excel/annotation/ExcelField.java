package com.fast.poi.excel.annotation;

import com.fast.poi.excel.constants.AlignType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel注解
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.common.annotation
 * @fileNmae Excel
 * @date 2023-8-29
 * @copyright
 * @since 0.0.1
 */
@Retention(value = RetentionPolicy.RUNTIME)
@Target(value = ElementType.FIELD)
public @interface ExcelField {

    /** 单元格标题 */
    String title() default "";

    /** 单元格主标题（多个相同的主标题会自动合并） */
    String mainTitle() default "";

    /** 单元格宽度（默认为20个字符的宽度） */
    int width() default 20 * 256;

    /** 对齐方式（默认左对齐） */
    AlignType align() default AlignType.LEFT;

    /** 当前列是否开启行合并 */
    boolean mergeRow() default false;

    /** 合并依赖（选择指定字段，作为合并依据，与指定字段合并保持一致）*/
    String mergeDepend() default "";

    /** 日期格式 */
    String format() default "";

    /** 背景色，默认白色 */
    IndexedColors bgColor() default IndexedColors.WHITE;

}
