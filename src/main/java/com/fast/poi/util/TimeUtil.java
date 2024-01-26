package com.fast.poi.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.util
 * @fileNmae TimeUtil
 * @date 2023-9-1
 * @copyright
 * @since 0.0.1
 */
public class TimeUtil {

    public static final SimpleDateFormat STANDARD_DATETIME_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    public static final SimpleDateFormat STANDARD_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");
    public static final SimpleDateFormat STANDARD_TIME_FORMAT = new SimpleDateFormat("HH:mm:ss");

    public static Date getDate(String str, SimpleDateFormat format) {
        format = format == null ? STANDARD_DATETIME_FORMAT : format;
        try {
            return format.parse(str);
        } catch (ParseException e) {
            return null;
        }
    }

    public static long getTimestamp(String str, SimpleDateFormat format) {
        Date date = getDate(str, format);
        if (date != null) {
            return date.getTime();
        }
        return 0;
    }

    public static long getTimestamp(String str){
        return getTimestamp(str, null);
    }

    /**
     * 获取指定格式的日期字符串
     * @param date 日期
     * @param format 日期格式
     * @return
     */
    public static String getTimeStr(Date date, String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }
}
