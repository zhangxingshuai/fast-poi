package com.fast.poi.word.service;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * 文档循环处理
 *
 * @author 张兴帅
 * @projectName aidustry-boot
 * @packageNmae fast.poi.word.service
 * @fileNmae ILoopService
 * @date 2023-9-7
 * @copyright
 * @since 0.0.1
 */
@Deprecated
public interface ILoopService {

    void resolveLoopRun(XWPFDocument doc);

    void resolveLoopParagraphs();

}
