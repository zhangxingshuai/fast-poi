package com.fast.poi.excel.bean;

import lombok.Data;

/**
 * excel下拉选对象
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.excel.bean
 * @fileNmae DropDownBox
 * @date 2023-9-1
 * @copyright
 * @since 0.0.1
 */
@Data
public class DropDownBox {

    private Boolean active;

    private String[] values;

    private String[] valuesMap;

}
