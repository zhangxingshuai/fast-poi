package com.fast.poi.word.bean;

import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.ArrayList;
import java.util.List;

/**
 * word 上下文对象
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae WordContextBean
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Data
public class WordContextBean {

    /** 文档对象 */
    XWPFDocument doc;

    /** 导出数据源对象 */
    Object datasource;

    /** 文本类型字段 */
    List<WordTextBean> textBeanList;

    /** 表格类型字段 */
    List<WordTableBean> tableBeanList;

    /** 循环配置 */
    List<WordLoopBean> loopBeanList;

    public void addTextBean(WordTextBean textBean){
        if (textBeanList == null) {
            textBeanList = new ArrayList<>();
        }
        textBeanList.add(textBean);
    }

    public void addTableBean(WordTableBean tableBean){
        if (tableBeanList == null) {
            tableBeanList = new ArrayList<>();
        }
        tableBeanList.add(tableBean);
    }


    public void addWordLoop(WordLoopBean wordLoopBean) {
        if (loopBeanList == null) {
            loopBeanList = new ArrayList<>();
        }
        loopBeanList.add(wordLoopBean);
    }

    public void setDoc(XWPFDocument doc){
        this.doc = doc;
    }

    public void setDatasource(Object datasource) {
        this.datasource = datasource;
    }
}
