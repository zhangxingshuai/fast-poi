package com.fast.poi.word.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * word 文本字段注解，表示字段是word文本类型
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.annotation
 * @fileNmae WordText
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface WordText {

    /** word字段名称*/
    String name() default "";
}
