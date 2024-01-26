package com.fast.poi.word.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * word
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.annotation
 * @fileNmae WordIterator
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface WordLoop {

    String key() default "";
}
