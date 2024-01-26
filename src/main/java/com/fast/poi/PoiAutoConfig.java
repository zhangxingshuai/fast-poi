package com.fast.poi;

import com.fast.poi.excel.service.IExcelService;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.context.annotation.ComponentScan;

/**
 *
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi
 * @fileNmae AutoConfig
 * @date 2023-8-31
 * @copyright
 * @since 0.0.1
 */
//@EnableAutoConfiguration
@ComponentScan
@ConditionalOnClass({IExcelService.class})
public class PoiAutoConfig {
}
