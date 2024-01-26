package com.fast.poi.word.example.vo;

import com.fast.poi.word.annotation.WordLoop;
import com.fast.poi.word.annotation.WordTable;
import com.fast.poi.word.annotation.WordText;
import lombok.Data;

import java.util.List;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example.vo
 * @fileNmae ExportVO
 * @date 2023-9-4
 * @copyright
 * @since
 */
@Data
public class ExportVO {

    @WordText
    private String title;

    @WordText
    private String year;

    @WordLoop(key = "articles")
    private List<Article> articleList;

    @WordTable
    private List<Article> articleTable;

    @WordText
    private Integer count;

    @WordText
    private Integer score;
}
