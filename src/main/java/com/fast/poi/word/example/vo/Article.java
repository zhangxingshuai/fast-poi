package com.fast.poi.word.example.vo;

import lombok.Data;

/**
 * 这个文件是干什么用的
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.example.vo
 * @fileNmae Article
 * @date 2023-9-4
 * @copyright
 * @since
 */
@Data
public class Article implements Cloneable {

    private Integer index;

    private String title;

    private String author;

    private String content;
}
