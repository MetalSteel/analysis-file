package com.lanlinker.service;

import org.junit.Test;

import static org.junit.Assert.*;

/**
 * Created by wanggang on 2017/7/21.
 */
public class Word2HtmlServiceTest {
    @Test
    public void doc2Html() throws Exception {
        Word2HtmlService word2HtmlService = new Word2HtmlService();
        word2HtmlService.doc2Html("什么是云_mac.doc");
    }

    @Test
    public void docx2Html() throws Exception {
        Word2HtmlService word2HtmlService = new Word2HtmlService();
        word2HtmlService.docx2Html("springboot_笔记.docx");
    }

}