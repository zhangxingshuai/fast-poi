package com.fast.poi.util;

import com.fast.poi.word.constants.Placeholder;
import org.apache.poi.xwpf.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * word 工具类
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.util
 * @fileNmae WordUtil
 * @date 2023-9-4
 * @copyright 
 * @since 0.0.1
 */
public class WordUtil {

    private static final Pattern PATTERN_TEXT = Pattern.compile(Placeholder.START.getValue() + ".*?" + Placeholder.END.getValue());
    private static final Pattern PATTERN_LOOP_TEXT = Pattern.compile(Placeholder.LOOP_TEXT_START.getValue() + ".*?" + Placeholder.LOOP_TEXT_END.getValue());

    private static final Pattern PATTERN_LOOP = Pattern.compile(Placeholder.LOOP_START.getValue() + ".*?" + Placeholder.LOOP_END.getValue());

    /** 获取当前文本的占位符列表 */
    public static List<String> getTextName(String text) {
        List<String> nameList = new ArrayList<>();
        if (text == null) {
            return nameList;
        }
        Matcher matcher = PATTERN_TEXT.matcher(text);
        while (matcher.find()) {
            nameList.add(matcher.group().replaceAll(Placeholder.START.getValue(), "").replaceAll(Placeholder.END.getValue(), ""));
        }
        return nameList;
    }

    public static List<String> getLoopTextName(String text) {
        List<String> nameList = new ArrayList<>();
        Matcher matcher = PATTERN_LOOP_TEXT.matcher(text);
        while (matcher.find()) {
            nameList.add(matcher.group().replaceAll(Placeholder.LOOP_TEXT_START.getValue(), "").replaceAll(Placeholder.LOOP_TEXT_END.getValue(), ""));
        }
        return nameList;
    }

    public static void main(String[] args) {
        Matcher matcher = PATTERN_LOOP_TEXT.matcher("123#{hello}456");
        while (matcher.find()) {
            System.out.println(matcher.group());
        }
    }

    /** 获取table列的占位符列表 */
    public static List<String> getTableColumns(XWPFTable xwpfTable) {
        List<String> nameList = new ArrayList<>();
        int rows = xwpfTable.getNumberOfRows();
        if (rows < 1) {
            return nameList;
        }
        XWPFTableRow row = xwpfTable.getRow(1);
        List<ICell> cells = row.getTableICells();
        for (ICell cell : cells) {
            String text = ((XWPFTableCell) cell).getText();
            Matcher matcher = PATTERN_TEXT.matcher(text);
            while (matcher.find()) {
                nameList.add(matcher.group().replaceAll(Placeholder.START.getValue(), "").replaceAll(Placeholder.END.getValue(), ""));
            }
        }
        // 移除占位符所在行
//        xwpfTable.removeRow(1);
        return nameList;
    }


    public static String getLoopText(String str) {
        Matcher matcher = PATTERN_LOOP.matcher(str);
        if (matcher.find()) {
            return matcher.group();
        }
        return "";
    }


    public static void copyRunStyle(XWPFRun oldRun, XWPFRun newRun) {
        newRun.setStyle(oldRun.getStyle());
        newRun.setFontFamily(oldRun.getFontFamily());
        if (oldRun.getFontSize() != -1) {
            newRun.setFontSize(oldRun.getFontSize());
        }
    }
}
