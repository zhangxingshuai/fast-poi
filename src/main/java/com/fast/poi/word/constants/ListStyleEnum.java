package com.fast.poi.word.constants;

import com.deepoove.poi.data.NumberingFormat;
import lombok.Getter;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.constants
 * @fileNmae ListStyle
 * @date 2023-9-11
 * @copyright
 * @since 0.0.1
 */
@Getter
public enum ListStyleEnum {
    BULLET(NumberingFormat.BULLET), // ●

    DECIMAL(NumberingFormat.DECIMAL), // 1. 2. 3.

    DECIMAL_PARENTHESES(NumberingFormat.DECIMAL_PARENTHESES), // 1) 2) 3)

    LOWER_LETTER(NumberingFormat.LOWER_LETTER), // a. b. c.

    LOWER_ROMAN(NumberingFormat.LOWER_ROMAN), // i ⅱ ⅲ

    UPPER_LETTER(NumberingFormat.UPPER_LETTER), // A. B. C.

    UPPER_ROMAN(NumberingFormat.UPPER_ROMAN); // Ⅰ Ⅱ Ⅲ

    private NumberingFormat format;
    ListStyleEnum(NumberingFormat format) {
        this.format = format;
    }
}
