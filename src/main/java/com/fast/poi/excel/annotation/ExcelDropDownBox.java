package com.fast.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel单元格下拉框注解
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.annotation
 * @fileNmae DropDownBox
 * @date 2023-9-1
 * @copyright
 * @since 0.0.1
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelDropDownBox {

    /** 下拉选项数值 */
    String[] values() default {};

    /** 数值与展示项的映射 */
    String[] valuesMap() default {};

    /** 激活状态，默认激活，只有激活才生效*/
    boolean active() default true;

}
