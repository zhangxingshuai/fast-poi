package com.fast.poi.word.bean;

import com.fast.poi.word.annotation.WordLoop;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.List;

/**
 * word
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae WordListBean
 * @date 2023-9-5
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Data
public class WordLoopBean {

    private String key;

    private List<Object> loopData;

    public WordLoopBean(Field field, Object obj) throws IllegalAccessException {
        WordLoop annotation = field.getAnnotation(WordLoop.class);
        String k = annotation.key();
        if ("".equals(k)) {
            k = field.getName();
        }
        key = k;
        loopData = (List) field.get(obj);
    }
}
