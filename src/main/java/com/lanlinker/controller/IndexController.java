package com.lanlinker.controller;

import com.lanlinker.service.Word2HtmlService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

/**
 * 首页控制器
 * Created by wanggang on 2017/7/21.
 * @author wanggang
 * @version 1.0
 */
@Controller
public class IndexController {
    /**
     * 注入业务层
     */
    @Autowired
    private Word2HtmlService word2HtmlService;

    /**
     * 访问首页
     * @return
     */
    @GetMapping("/index")
    public String index(){
        return "/index";
    }

    /**
     * doc文件转换为html文件
     * @return
     */
    @GetMapping("/doc2Html")
    public String doc2Html() throws ParserConfigurationException, TransformerException, IOException {
        word2HtmlService.doc2Html("http://localhost:8081/springboot_%E7%AC%94%E8%AE%B0/springboot_%E7%AC%94%E8%AE%B0.html");
        return "";
    }

}
