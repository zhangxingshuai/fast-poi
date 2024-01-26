package com.fast.poi.word.example.vo;


import com.fast.poi.util.NumberUtil;
import com.fast.poi.word.annotation.*;
import lombok.Data;
import org.springframework.beans.BeanUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example.vo
 * @fileNmae DuibiaoExportVO
 * @date 2023-9-7
 * @copyright 
 * @since
 */
@Data
@SpringEL
public class DuibiaoExportVO {

    @WordText
    private String orgname;

    @WordText
    private Integer dbnd;

    @WordText
    private String kzrq;

    @WordText
    private Integer cnt;
    @WordText
    private Double zhfpf = 0D;

    @WordLoop(key = "dbList")
    @WordTable(index = 0)
    @LoopBlock
    private List<DbItem> dbList;

    @WordLoop(key = "goodItems")
    private List<DbItem> goodItems;

    public void setDbList(List<DbItem> dbList) {
        this.dbList = dbList;
        this.goodItems = new ArrayList<>();
        for (DbItem dbItem : dbList) {
            if (dbItem.getFpfs() > dbItem.getBzfs()*0.8) {
                DbItem item = new DbItem();
                BeanUtils.copyProperties(dbItem, item);
                item.setXh(NumberUtil.getChineseNum(item.getXh()));
                this.goodItems.add(item);
                this.zhfpf += item.getFpfs();
            }
        }
    }

}
