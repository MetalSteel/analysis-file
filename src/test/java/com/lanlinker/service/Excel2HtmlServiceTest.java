package com.lanlinker.service;

import org.junit.Test;

/**
 * Created by wanggang on 2017/7/22.
 */
public class Excel2HtmlServiceTest {
    @Test
    public void excel2Html() throws Exception {
        Excel2HtmlService excel2HtmlService = new Excel2HtmlService();
        excel2HtmlService.excel2Html("文库.xls");
    }

}