package com.fast.poi.word.bean;

import com.deepoove.poi.data.NumberingFormat;
import lombok.Data;

/**
 * 列表项样式
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae ListStyle
 * @date 2023-9-11
 * @copyright
 * @since 0.0.1
 */
@Data
public class ListStyle {

    /** 列表项前缀 */
    private NumberingFormat prefix;

    /** 左侧缩进字符个数 */
    private Double indentLeftChars;
}
