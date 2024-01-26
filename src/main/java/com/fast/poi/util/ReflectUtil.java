package com.fast.poi.util;

import java.lang.reflect.Field;

/**
 * 反射工具类
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.util
 * @fileNmae ReflectUtil
 * @date 2023-9-5
 * @copyright
 * @since 0.0.1
 */
public class ReflectUtil {

    public static Object getValue(Object obj, String fieldName){
        Class<?> clazz = obj.getClass();
        try {
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            return field.get(obj);
        } catch (Exception e) {
        }
        return null;
    }
}
