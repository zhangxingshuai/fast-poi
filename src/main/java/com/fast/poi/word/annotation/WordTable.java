package com.fast.poi.word.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.annotation
 * @fileNmae Table
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface WordTable {

    /** word表格索引，在word文档中第几个table*/
    int index() default 0;

}
