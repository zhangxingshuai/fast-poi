package com.fast.poi.util;

import java.util.HashMap;
import java.util.Map;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.util
 * @fileNmae NumberUtil
 * @date 2023-9-7
 * @copyright
 * @since 0.0.1
 */
public class NumberUtil {

    private static final Map<String,String> convertMap = new HashMap<>();
    static {
        convertMap.put("1","一");
        convertMap.put("2","二");
        convertMap.put("3","三");
        convertMap.put("4","四");
        convertMap.put("5","五");
        convertMap.put("6","六");
        convertMap.put("7","七");
        convertMap.put("8","八");
        convertMap.put("9","九");
        convertMap.put("10","十");
    }


    public static String getChineseNum(String num){
        return convertMap.get(num);
    }
}
