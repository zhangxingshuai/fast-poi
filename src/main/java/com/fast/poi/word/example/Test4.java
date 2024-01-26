package com.fast.poi.word.example;

import com.fast.poi.word.example.vo.DbItem;
import com.fast.poi.word.example.vo.DuibiaoExportVO;
import com.fast.poi.word.service.impl.WordServiceImpl;
import com.deepoove.poi.XWPFTemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.wordx
 * @fileNmae Test
 * @date 2023-9-8
 * @copyright
 * @since
 */
public class Test4 {

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
            dbItem.setRate(dbItem.getFpfs()/ dbItem.getBzfs()*100+"%");
            dbList.add(dbItem);
        }
        exportVO.setDbList(dbList);
        exportVO.setCnt(dbList.size());

        try {
            XWPFTemplate template = new WordServiceImpl().render("template/word/对标报告2.docx", exportVO);
            template.writeAndClose(new FileOutputStream(new File("对标报告4.docx")));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
