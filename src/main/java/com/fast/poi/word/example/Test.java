package com.fast.poi.word.example;

import com.fast.poi.word.example.vo.Article;
import com.fast.poi.word.example.vo.ExportVO;
import com.fast.poi.word.service.IWordService;
import com.fast.poi.word.service.impl.WordServiceImpl;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example
 * @fileNmae Test
 * @date 2023-9-4
 * @copyright
 * @since
 */
public class Test {

    public static void main(String[] args) {
        ExportVO exportVO = new ExportVO();
        exportVO.setTitle("文章精选");
        exportVO.setYear("2023");
        List<Article> articles = new ArrayList<>(10);
        List<Article> articleTable = new ArrayList<>(10);
        for (int i = 0; i < 10; i++) {
            Article article = new Article();
            article.setIndex(i+1);
            article.setTitle("文章"+(i+1));
            article.setAuthor("作者"+i+1);
            article.setContent("文章内容.............................");
            articles.add(article);
            articleTable.add(article);
        }
        exportVO.setArticleList(articles);
        exportVO.setArticleTable(articleTable);
        exportVO.setCount(articles.size());
        exportVO.setScore(articles.size()*10);
        IWordService wordService = new WordServiceImpl();
        InputStream ins = Test.class.getClassLoader().getResourceAsStream("template/word/模板1.docx");
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(ins);
            wordService.export(doc,exportVO);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
