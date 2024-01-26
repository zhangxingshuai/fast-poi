package com.fast.poi.word.example.vo;

import com.fast.poi.word.bean.ListData;
import com.fast.poi.word.bean.ListStyle;
import com.deepoove.poi.data.NumberingFormat;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example.vo
 * @fileNmae DbItem
 * @date 2023-9-7
 * @copyright
 * @since
 */
@Data
public class DbItem {

    private String xh;
    private String flname;
    private Integer bzfs;
    private Double fpfs;
    private String rate;

    private ListData itemList = new ListData(NumberingFormat.BULLET);

    public void setFpfs(Double fpfs) {
        this.fpfs = fpfs;
        if (this.fpfs >= 0.6 * bzfs) {
            this.itemList = new ListData("项目1","项目2","项目3");
        }
        ListStyle listStyle = new ListStyle();
        listStyle.setPrefix(NumberingFormat.BULLET);
        listStyle.setIndentLeftChars(4D);
        this.itemList.setListStyle(listStyle);
    }

    @Data
    @AllArgsConstructor
    public  static class Item {
        private String name;
    }
}
