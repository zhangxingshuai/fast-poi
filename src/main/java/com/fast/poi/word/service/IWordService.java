package com.fast.poi.word.service;

import com.deepoove.poi.XWPFTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * word相关接口
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae com.fast.poi.word.service
 * @fileNmae IWordService
 * @date 2023-9-4
 * @copyright
 * @since 0.0.1
 */
public interface IWordService {

    @Deprecated
    void export(XWPFDocument doc, Object exportVO);


    XWPFTemplate render(String relativePath, Object exportVO);

}
