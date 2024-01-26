package com.fast.poi.word.example;

import com.fast.poi.util.NumberUtil;
import com.fast.poi.word.example.vo.DbqkVO;
import com.fast.poi.word.service.impl.WordServiceImpl;
import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.wordx
 * @fileNmae Test2
 * @date 2023-9-11
 * @copyright
 * @since
 */
public class Test3 {

    public static void main(String[] args) {
        DbqkVO dbqkVO = new DbqkVO();
        dbqkVO.setDbnd("2022");
        dbqkVO.setOrgname("xx公司");
        List<DbqkVO.Category> categoryList = new ArrayList<>(5);
        for (int i = 0; i < 5; i++) {
            DbqkVO.Category category = new DbqkVO.Category();
            category.setXh(NumberUtil.getChineseNum(i+1+""));
            category.setFlname("分类"+(i+1));
            category.setSumbzfs(10D);
            category.setSumfpfs(i+1D);
            List<DbqkVO.Item> itemList = new ArrayList<>(3);
            for (int j = 0; j < 3; j++) {
                DbqkVO.Item item = new DbqkVO.Item();
                item.setXh(String.valueOf(j+1));
                item.setZbxname("指标项"+(j+1));
                item.setSumfpfs(i+1D);
                List<DbqkVO.Jcnr> jcnrList = new ArrayList<>();
                for (int k = 0; k < 2; k++) {
                    DbqkVO.Jcnr jcnr = new DbqkVO.Jcnr();
                    jcnr.setXh((j+1)+"."+(k+1));
                    jcnr.setJcnr("检查内容"+(k+1));
                    jcnrList.add(jcnr);
                }
                item.setJcnrList(jcnrList);
                itemList.add(item);
            }
            category.setItemList(itemList);
            categoryList.add(category);
        }
        dbqkVO.setCategoryList(categoryList);
        try {
            XWPFTemplate template = new WordServiceImpl().render("template/word/对标情况.docx", dbqkVO);
            template.writeToFile("对标说明情况导出.docx");
            template.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
