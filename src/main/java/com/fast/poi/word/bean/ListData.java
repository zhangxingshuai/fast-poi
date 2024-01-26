package com.fast.poi.word.bean;

import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.ParagraphStyle;
import org.springframework.beans.BeanUtils;

import java.util.Arrays;
import java.util.List;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae ListData
 * @date 2023-9-11
 * @copyright
 * @since
 */
public class ListData extends NumberingRenderData {

    public ListData(NumberingFormat format) {
        super.setFormats(Arrays.asList(format));
    }

    public ListData(String ...items) {
        NumberingRenderData data = Numberings.of(items).create();
        super.setItems(data.getItems());
    }

    /** 新增列表项 */
    public void addItem(String item){
        List<NumberingItemRenderData> items = Numberings.of(item).create().getItems();
        int size = super.getItems().size();
        if (size > 0) {
            // 设置新增列表项的样式，与之前样式保持一致
            ParagraphStyle style = super.getItems().get(0).getItem().getParagraphStyle();
            items.get(0).getItem().setParagraphStyle(style);
        }
        super.getItems().addAll(items);
    }

    /** 设置列表项样式 */
    public void setListStyle(ListStyle style) {
        // 设置段落前缀样式
        if (style.getPrefix() != null) {
            super.setFormats(Arrays.asList(style.getPrefix()));
        }
        ParagraphStyle paragraphStyle = new ParagraphStyle();
        BeanUtils.copyProperties(style, paragraphStyle);
        for (NumberingItemRenderData item : super.getItems()) {
            item.getItem().setParagraphStyle(paragraphStyle);
        }
    }



}
