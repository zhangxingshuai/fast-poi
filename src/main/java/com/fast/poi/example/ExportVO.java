package com.fast.poi.example;

import com.fast.poi.excel.annotation.ExcelConditionStyle;
import com.fast.poi.excel.annotation.ExcelField;
import com.fast.poi.excel.annotation.ExcelGlobalStyle;
import com.fast.poi.excel.constants.AlignType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 导出测试类 数据位置： /resource/mock/data.txt
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.test.vo
 * @fileNmae ExportVO
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */

@ExcelGlobalStyle(showBorder = true, fontSize = 10, bold = true, fontName = "宋体")
public class ExportVO {

    @ExcelField(title = "序号", mergeRow = true, width = 8*256, align = AlignType.CENTER)
    @ExcelConditionStyle(condition = "categoryName = '领导重视'", fontColor = IndexedColors.RED,bold = true)
    private Integer index;

    @ExcelField(title = "分类名称(总:100分)", mergeRow = true,width = 40*256, align = AlignType.CENTER)
    @ExcelConditionStyle(condition = "categoryName = '领导重视'", bgColor = IndexedColors.LIGHT_YELLOW)
    private String categoryName;

    @ExcelField(title = "指标项名称", mergeRow = true,width = 40*256)
    private String itemName;

    @ExcelField(title = "检查内容",width = 40*256)
    private String content;

    @ExcelField(title = "评分标准",width = 40*256)
    private String standard;

    @ExcelField(title = "标准分数",width = 12*256, align = AlignType.RIGHT)
    @ExcelConditionStyle(condition = "score = 2", bgColor = IndexedColors.RED)
    private Double score;

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getCategoryName() {
        return categoryName;
    }

    public void setCategoryName(String categoryName) {
        this.categoryName = categoryName;
    }

    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getStandard() {
        return standard;
    }

    public void setStandard(String standard) {
        this.standard = standard;
    }

    public Double getScore() {
        return score;
    }

    public void setScore(Double score) {
        this.score = score;
    }
}
