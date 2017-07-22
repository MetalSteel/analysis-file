package com.lanlinker.service;

import org.junit.Test;

/**
 * Created by wanggang on 2017/7/22.
 */
public class PPT2HtmlServiceTest {
    @Test
    public void pptx2Html() throws Exception {
        PPT2HtmlService ppt2HtmlService = new PPT2HtmlService();
        ppt2HtmlService.pptx2Html("圆的极简创意封面简约大气通用商务ppt模板.pptx","png");
    }

    @Test
    public void ppt2Html() throws Exception {
        PPT2HtmlService ppt2HtmlService = new PPT2HtmlService();
        ppt2HtmlService.ppt2Html("UML基础教程(内部使用教程).ppt","png");
    }

}