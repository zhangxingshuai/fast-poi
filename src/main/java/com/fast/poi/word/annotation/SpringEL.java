package com.fast.poi.word.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 启用springEL表达式注解
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.annotation
 * @fileNmae SpringEL
 * @date 2023-9-11
 * @copyright
 * @since 0.0.1
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface SpringEL {

}
