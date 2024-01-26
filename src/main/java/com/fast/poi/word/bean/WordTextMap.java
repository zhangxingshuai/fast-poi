package com.fast.poi.word.bean;

import lombok.Data;

/**
 * word 文本映射
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae WordTextMap
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Data
public class WordTextMap {

    private String key;

    private String fieldName;

    private Object value;
}
