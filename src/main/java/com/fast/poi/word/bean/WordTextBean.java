package com.fast.poi.word.bean;

import com.fast.poi.word.annotation.WordText;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.lang.reflect.Field;

/**
 *
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae WordTextBean
 * @date 2023-9-5
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Data
@AllArgsConstructor
public class WordTextBean {

    /** 文本类型占位符名称 */
    private String name;

    /** 需要替换的值 */
    private Object value;

    public WordTextBean(Field field, Object obj) throws IllegalAccessException {
        WordText annotation = field.getAnnotation(WordText.class);
        String name = field.getName();
        if (!"".equals(annotation.name())) {
            name = annotation.name();
        }
        this.name = name;
        this.value = field.get(obj);
    }
}
