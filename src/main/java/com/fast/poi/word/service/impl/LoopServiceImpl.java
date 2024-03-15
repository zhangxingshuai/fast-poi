package com.fast.poi.word.service.impl;

import com.fast.poi.util.ReflectUtil;
import com.fast.poi.word.bean.WordLoopBean;
import com.fast.poi.word.constants.Placeholder;
import com.fast.poi.word.context.WordContext;
import com.fast.poi.word.service.ILoopService;
import com.fast.poi.util.WordUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 循环处理
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.service.impl
 * @fileNmae LoopServiceImpl
 * @date 2023-9-7
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Service
public class LoopServiceImpl implements ILoopService {

    @Override
    public void resolveLoopRun(XWPFDocument doc) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        List<XWPFRun> loopRun = new ArrayList<>();
        for (XWPFParagraph paragraph : paragraphs) {
            String start = "";
            String end = "";
            for (int i = 0; i < paragraph.getRuns().size(); i++) {
                XWPFRun run = paragraph.getRuns().get(i);
                String text = run.getText(0);
                if (text.contains(Placeholder.LOOP_START.getValue()) && text.contains(Placeholder.LOOP_END.getValue())) {
                    String loopText = WordUtil.getLoopText(text);
                    start = text.split("<LOOP:")[0];
                    end = text.split("</LOOP>")[1];
                    // 解析出数据源名称，循环体内占位符字段名
                    String content = resolveLoopText(loopText);
                    String newText = start + content.toString() + end;
                    XWPFRun newRun = paragraph.insertNewRun(i);
                    newRun.setText(newText);
                    WordUtil.copyRunStyle(run,newRun);
                    paragraph.removeRun(i+1);
                }

            }
        }
    }

    private String resolveLoopText(String loopText) {
        StringBuilder sb = new StringBuilder();
        List<String> fieldNames = WordUtil.getLoopTextName(loopText);
        Pattern pattern = Pattern.compile(":.*?\\>");
        Matcher matcher = pattern.matcher(loopText);
        if (matcher.find()) {
            String key = matcher.group().replaceAll(":","").replaceAll("\\>","");
            WordLoopBean wordLoopBean = WordContext.getWordLoopList().stream().filter(item -> item.getKey().equals(key)).findFirst().get();
            List<Object> loopData = wordLoopBean.getLoopData();
            String loopContent = loopText.replaceAll("<LOOP:"+key+">","").replaceAll("</LOOP>","");
            for (Object data : loopData) {
                String text = loopContent;
                for (String fieldName : fieldNames) {
                    String value = ReflectUtil.getValue(data, fieldName) + "";
                    text = text.replaceAll(Placeholder.LOOP_TEXT_START.getValue() + fieldName + Placeholder.LOOP_TEXT_END.getValue(), value);
                }
                sb.append(text);
            }
        }
        return sb.toString();
    }

    @Override
    public void resolveLoopParagraphs() {

    }

}
