package com.fast.poi.word.service.impl;

import com.fast.poi.util.ReflectUtil;

import com.fast.poi.word.bean.WordContextBean;
import com.fast.poi.word.bean.WordLoopBean;
import com.fast.poi.word.bean.WordTableBean;
import com.fast.poi.word.bean.WordTextBean;
import com.fast.poi.word.constants.Placeholder;
import com.fast.poi.word.context.WordContext;
import com.fast.poi.word.service.ILoopService;
import com.fast.poi.word.service.IWordService;
import com.fast.poi.util.WordUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.fast.poi.word.annotation.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

/**
 * word fast
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.service.impl
 * @fileNmae WordServiceImpl
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Slf4j
@Service
public class WordServiceImpl implements IWordService {

    @Resource
    private ILoopService loopService = new LoopServiceImpl();

    @Deprecated
    @Override
    public void export(XWPFDocument doc, Object exportVO) {
        // 解析模板
        resolveTemplate(doc, exportVO);
        long time = System.currentTimeMillis();
        try {
            doc.write(new FileOutputStream(new File("doc"+time+".docx")));
            doc.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /** 渲染模板
     * @param relativePath 模板相对路径
     * @param exportVO 导出对象
     * */
    @Override
    public XWPFTemplate render(String relativePath, Object exportVO) {
        if (relativePath == null || "".equals(relativePath)) {
            return null;
        }
        try {
            // 获取模板
            XWPFDocument template = new XWPFDocument(this.getClass().getClassLoader().getResourceAsStream(relativePath));
            // 解析配置
            Configure config = resolveConfig(exportVO);
            // 编译模板
            XWPFTemplate xwpfTemplate = XWPFTemplate.compile(template, config);
            // 渲染模板
            xwpfTemplate.render(exportVO);
            // 返回渲染后文档
            return xwpfTemplate;
        } catch (IOException e) {
            log.error("渲染模板异常-模板相对路径={}", relativePath);
            throw new RuntimeException(e);
        }
    }

    /** 解析导出配置 */
    private Configure resolveConfig(Object obj) {
        ConfigureBuilder builder = Configure.builder();
        if (obj.getClass().isAnnotationPresent(SpringEL.class)) {
            // 开启springEL
            builder.useSpringEL();
        }
        // 区块循环配置
        LoopRowTableRenderPolicy policy = new LoopRowTableRenderPolicy();
        for (Field field : obj.getClass().getDeclaredFields()) {
            if (field.isAnnotationPresent(LoopBlock.class)) {
                builder.bind(field.getName(), policy);
            }
        }
        return builder.build();
    }

    private void resolveTemplate(XWPFDocument doc, Object exportVO) {
        WordContextBean wordContextBean = new WordContextBean();
        wordContextBean.setDoc(doc);
        wordContextBean.setDatasource(exportVO);
        Class clazz = exportVO.getClass();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            try {
                field.setAccessible(true);
                if (field.isAnnotationPresent(WordText.class)){ // 普通文本
                    wordContextBean.addTextBean(new WordTextBean(field,exportVO));
                }
                if (field.isAnnotationPresent(WordTable.class)) { // 表格
                    wordContextBean.addTableBean(new WordTableBean(field, exportVO));
                }
                if (field.isAnnotationPresent(WordLoop.class)) { // 循环体
                    wordContextBean.addWordLoop(new WordLoopBean(field, exportVO));
                }
            } catch (Exception e) {
            }
        }
        WordContext.setWordContextBean(wordContextBean);
        // 解析普通文本
        resolveText(doc);
        // 解析表格
        resolveTable(doc);
        // 解析loop循环体
        resolveLoop(doc);
        loopService.resolveLoopRun(doc);
        // 清除word导出上下文
        WordContext.clear();
    }


    private void resolveTable(XWPFDocument doc) {
        List<XWPFTable> tables = doc.getTables();
        if (tables == null || tables.size() == 0) {
            return;
        }
        List<WordTableBean> tableList = WordContext.getWordTableList();
        for (WordTableBean wordTableBean : tableList) {
            Integer index = wordTableBean.getIndex();
            if (tables.size() <= index) {
                continue;
            }
            XWPFTable xwpfTable = tables.get(index);
            // 解析并替换固定位置占位符
            resolveTableFixedCell(xwpfTable);
            // 解析table列表项
            List<String> nameList = WordUtil.getTableColumns(xwpfTable);
            // 设置表格数据
            setTableData(xwpfTable,nameList,wordTableBean.getTableData());
        }
    }

    private void resolveTableFixedCell(XWPFTable xwpfTable) {
        List<WordTextBean> textBeanList = WordContext.getTextBeanList();
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (int i = 2; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            for (XWPFTableCell cell : row.getTableCells()) {
                XWPFParagraph paragraph = cell.getParagraphArray(0);
                String text = paragraph.getText();
                List<String> nameList = WordUtil.getTextName(text);
                if (nameList.size() > 0) {
                    for (String name : nameList) {
                        String value = textBeanList.stream().filter(item->item.getName().equals(name)).findFirst().get().getValue()+"";
                        text = text.replaceAll(Placeholder.START.getValue()+name+Placeholder.END.getValue(),value);
                    }
                    XWPFRun newRun = paragraph.createRun();
                    newRun.setText(text);
                    WordUtil.copyRunStyle(paragraph.getRuns().get(0),newRun);
                    paragraph.removeRun(0);
                }
            }
        }
    }

    /**设置表格数据
     * @param xwpfTable word表格对象
     * @param nameList 列名集合
     * @tableData 表格数据
     * */
    private void setTableData(XWPFTable xwpfTable, List<String> nameList, List<Object> tableData) {
        if (tableData == null || tableData.size() == 0) {
            return;
        }
        XWPFTableRow templateRow = xwpfTable.getRow(1);
        for (int i=0; i<tableData.size(); i++) {
            Object data = tableData.get(i);
            XWPFTableRow row = xwpfTable.insertNewTableRow(i+1);
            for (int j=0; j<nameList.size(); j++) {
                Object value = ReflectUtil.getValue(data, nameList.get(j));
                XWPFParagraph templateParagarph = templateRow.getCell(j).getParagraphArray(0);
                XWPFRun templateRun = templateParagarph.getRuns().get(0);
                XWPFTableCell cell = row.addNewTableCell();
                XWPFParagraph paragraph = cell.getParagraphArray(0);
                paragraph.setStyle(templateParagarph.getStyle());
                paragraph.setAlignment(templateParagarph.getAlignment());
                XWPFRun run = paragraph.createRun();
                run.setText(value+"");
                WordUtil.copyRunStyle(templateRun,run);
            }
        }
        xwpfTable.removeRow(tableData.size()+1);
    }

    private void resolveLoop(XWPFDocument doc) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        String key = "";
        boolean match = false;
        Map<String,List<XWPFParagraph>> loopMap = new HashMap<>();
        List<XWPFParagraph> paragraphList = null;
        int loopStart = 0;
        int loopEnd = 0;
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                if (match && !paragraphList.contains(paragraph) && !run.getText(0).equals(Placeholder.LOOP_END.getValue())) {
                    paragraphList.add(paragraph);
                }
                if (run.getText(0).startsWith(Placeholder.LOOP_START.getValue())) { // LIST 开始标识
                    loopStart++;
                    if (loopStart>1) {
                        continue;
                    }
                    match = true;
                    key = run.getText(0).replaceAll(Placeholder.LOOP_START.getValue(), "").replaceAll(">","");
                    paragraphList = new ArrayList<>();
                    paragraph.removeRun(i);
                } else if (run.getText(0).equals(Placeholder.LOOP_END.getValue())) { // LIST结束标识
                    loopEnd++;
                    if (loopEnd == loopStart) {
                        match = false;
                        loopMap.put(key, paragraphList);
                        paragraph.removeRun(i);
                    } else if (loopEnd < loopStart) {
                        paragraphList.add(paragraph);
                    }
                }
            }
        }
        setLoopData(loopMap);
    }

    private void setLoopData(Map<String, List<XWPFParagraph>> loopMap) {
        List<WordLoopBean> loopList = WordContext.getWordLoopList();
        if (loopList == null) { // 如果循环数据为空，清除文档中对应的模板
            removeTemplateParagarph(loopMap);
            return;
        }
        XWPFDocument doc = WordContext.getDoc();
        for (WordLoopBean wordLoopBean : loopList) {
            List<Object> loopData = wordLoopBean.getLoopData();
            List<XWPFParagraph> paragraphList = loopMap.get(wordLoopBean.getKey());
            for (Object data : loopData) {
                // 要插入段落位置游标（第一次从模板段落游标位置插入，之后从新增段落游标位置插入）
                XmlCursor cursor = null;
                // 当前新增段落
                XWPFParagraph newParagarph = null;
                if (paragraphList != null) {
                    for (int i=0; i<paragraphList.size(); i++) {
                        XWPFParagraph paragraph = paragraphList.get(i);
                        if (i==0) {
                            cursor = paragraph.getCTP().newCursor();
                        } else {
                            cursor = newParagarph.getCTP().newCursor();
                        }
                        newParagarph = doc.insertNewParagraph(cursor);
                        newParagarph.setStyle(paragraph.getStyle());
                        XWPFRun newRun = newParagarph.createRun();
                        for (XWPFRun run : paragraph.getRuns()) {
                            WordUtil.copyRunStyle(run,newRun);
                            // 判断当前段落是否包含循环
                            if (run.getText(0).startsWith(Placeholder.LOOP_START.getValue())) {

                            }
                            // 文本模板
                            List<String> nameList = WordUtil.getLoopTextName(run.getText(0));
                            if (nameList.size() > 0) { // 需要进行替换
                                replaceLoopText(newRun,run.getText(0),nameList,data);
                            } else { // 不需要替换，直接使用原文本
                                newRun.setText(run.getText(0));
                            }
                        }
                    }
                }
            }
        }
        removeTemplateParagarph(loopMap);
    }

    /** 移除模板段落 */
    private void removeTemplateParagarph(Map<String, List<XWPFParagraph>> loopMap) {
        Set<Map.Entry<String, List<XWPFParagraph>>> entries = loopMap.entrySet();
        for (Map.Entry<String, List<XWPFParagraph>> entry : entries) {
            List<XWPFParagraph> list = entry.getValue();
            for (XWPFParagraph paragraph : list) {
                for (int i = 0; i < paragraph.getRuns().size(); i++) {
                    paragraph.removeRun(i);
                }
            }
        }
    }


    /** 替换循环体内文本*/
    private void replaceLoopText(XWPFRun newRun, String text, List<String> nameList, Object data) {
        for (String name : nameList) {
            // 获取值
            String value = ReflectUtil.getValue(data, name)+"";
            // 替换值
            text = text.replaceAll(Placeholder.LOOP_TEXT_START.getValue()+name+Placeholder.LOOP_TEXT_END.getValue(), value);
        }
        newRun.setText(text);
    }

    /** 解析文本 */
    private void resolveText(XWPFDocument doc) {
        List<WordTextBean> textBeanList = WordContext.getTextBeanList();
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            List<XWPFRun> loopRuns = new ArrayList<>();
            for (int i=0; i<runs.size(); i++) {
                XWPFRun run = runs.get(i);
                List<String> nameList = WordUtil.getTextName(run.getText(0));
                if (nameList.size() > 0) { // 需要进行替换
                    // 创建新run
                    XWPFRun newRun = paragraph.insertNewRun(i);
                    // 替换旧文本并创建新段落
                    replaceText(run,newRun,nameList,textBeanList);
                    // 移除旧文本段落
                    paragraph.removeRun(i+1);
                }
            }
        }
    }

    private void replaceText(XWPFRun run, XWPFRun newRun, List<String> nameList, List<WordTextBean> textBeanList) {
        // 获取旧文本
        String text = run.getText(0);
        // 遍历占位符列表
        for (String name : nameList) {
            // 获取数据
            for (WordTextBean wordTextBean : textBeanList) {
                if (name.equals(wordTextBean.getName())) {
                    String value = String.valueOf(wordTextBean.getValue());
                    text = text.replaceAll(Placeholder.START.getValue()+name+Placeholder.END.getValue(), value);
                }
            }
        }
        // 创建新的文本段落
        newRun.setText(text);
        WordUtil.copyRunStyle(run, newRun);
    }

}
