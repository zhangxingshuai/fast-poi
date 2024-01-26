package com.fast.poi.word.example;

import com.fast.poi.word.example.vo.DbItem;
import com.fast.poi.word.example.vo.DuibiaoExportVO;
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
 * @fileNmae Test2
 * @date 2023-9-7
 * @copyright
 * @since 0.0.1
 */
public class Test2 {

    public static void main(String[] args) {
        DuibiaoExportVO exportVO = new DuibiaoExportVO();
        exportVO.setOrgname("信息公司");
        exportVO.setDbnd(2022);
        exportVO.setKzrq("2022年2月");
        List<DbItem> dbList = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
            DbItem dbItem = new DbItem();
            dbItem.setXh(String.valueOf(i+1));
            dbItem.setFlname("指标分类"+(i+1));
            dbItem.setBzfs(100);
            dbItem.setFpfs((i+1)*20D);
            dbItem.setRate(Math.ceil(dbItem.getFpfs()/ dbItem.getBzfs())+"%");
            dbList.add(dbItem);
        }
        exportVO.setDbList(dbList);
        exportVO.setCnt(dbList.size());
        IWordService wordService = new WordServiceImpl();
        InputStream ins = Test.class.getClassLoader().getResourceAsStream("template/word/模板2.docx");
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(ins);
            wordService.export(doc,exportVO);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
