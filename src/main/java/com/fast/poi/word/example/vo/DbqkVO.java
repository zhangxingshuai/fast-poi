package com.fast.poi.word.example.vo;

import com.fast.poi.word.annotation.LoopBlock;
import lombok.Data;

import java.util.List;

/**
 * 对标情况
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example.vo
 * @fileNmae DbqkVO
 * @date 2023-9-11
 * @copyright
 * @since
 */
@Data
public class DbqkVO {

    private String orgname;

    private String dbnd;

    @LoopBlock
    private List<Category> categoryList;

    @Data
    public static class  Category {

        private String xh;

        private String flname;

        private Double sumbzfs;

        private Double sumfpfs;

        private List<Item> itemList;

    }

    @Data
    public static class Item {
        private String xh;
        private String zbxname;
        private Double sumfpfs;
        private List<Jcnr> jcnrList;
    }

    @Data
    public static class Jcnr {

        private String xh;

        private String jcnr;
    }
}
