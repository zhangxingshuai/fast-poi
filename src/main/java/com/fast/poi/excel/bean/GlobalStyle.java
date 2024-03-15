package com.fast.poi.excel.bean;

import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * excel单元格全局样式
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.chng.hnxxgs.aidustry.poi.excel.bean
 * @fileNmae GlobalStyle
 * @date 2023-8-30
 * @copyright 华能信息技术有限公司
 * @since 0.0.1
 */
@Data
public class GlobalStyle extends BaseStyle {

    /** 是否显示边框 */
    private Boolean showBorder;

    /** 字体大小 */
    private Short fontSize;

    /** 字体名称 */
    private String fontName;

    /** 字体加粗 */
    private Boolean bold;

    /** 背景颜色 */
    public IndexedColors bgColor;

    /** 标题背景色 */
    public IndexedColors titleBgColor;

    /** 行高 */
    public Float rowHieght;

    /** 冻结标题 */
    public Boolean freezeTitle;
}
