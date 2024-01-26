package com.fast.poi.word.context;

import com.fast.poi.word.bean.WordContextBean;
import com.fast.poi.word.bean.WordLoopBean;
import com.fast.poi.word.bean.WordTableBean;
import com.fast.poi.word.bean.WordTextBean;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.ArrayList;
import java.util.List;

/**
 * word 上下文
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.context
 * @fileNmae WordContext
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
@Deprecated
public class WordContext {

    private static final ThreadLocal<WordContextBean> WORD_CONTEXT_BEAN = new ThreadLocal<>();

    private static WordContextBean get(){
        WordContextBean wordContextBean = WORD_CONTEXT_BEAN.get();
        wordContextBean = wordContextBean == null ? new WordContextBean() : wordContextBean;
        return wordContextBean;
    }


    public static void setWordContextBean(WordContextBean contextBean) {
        WORD_CONTEXT_BEAN.set(contextBean);
    }


    /** 获取导出文本列表 */
    public static List<WordTextBean> getTextBeanList() {
        List<WordTextBean> textBeanList = get().getTextBeanList();
        return textBeanList == null ? new ArrayList<>() : textBeanList;
    }

    /** 获取导出表格列表 */
    public static List<WordTableBean> getWordTableList() {
        List<WordTableBean> tableBeanList = get().getTableBeanList();
        return tableBeanList == null ? new ArrayList<>() : tableBeanList;
    }

    /** 获取导出循环列表 */
    public static List<WordLoopBean> getWordLoopList() {
        List<WordLoopBean> loopBeanList = get().getLoopBeanList();
        return loopBeanList == null ? new ArrayList<>() : loopBeanList;
    }

    public static XWPFDocument getDoc(){
        return get().getDoc();
    }

    public static Object getDatasource(){
        return WORD_CONTEXT_BEAN.get().getDatasource();
    }
    public static void clear(){
        WORD_CONTEXT_BEAN.remove();
    }
}
