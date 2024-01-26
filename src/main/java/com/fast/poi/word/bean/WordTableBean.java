package com.fast.poi.word.bean;

import com.fast.poi.word.annotation.WordTable;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.List;

/**
 *
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.bean
 * @fileNmae WordTableBean
 * @date 2023-9-5
 * @copyright
 * @since 0.0.1
 */
@Deprecated
@Data
public class WordTableBean {

    /** 表格在文档中的索引 */
    private Integer index;

    /** table表格数据 */
    private List<Object> tableData;

    public WordTableBean(Field field, Object obj) throws Exception {
        WordTable annotation = field.getAnnotation(WordTable.class);
        this.index = annotation.index();
        this.tableData = (List) field.get(obj);
    }

}
