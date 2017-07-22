package com.lanlinker.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

/**
 * Word文档转换成html
 * Created by wanggang on 2017/7/21.
 * @author wanggang
 * @version 1.0
 */
@Service
public class Word2HtmlService {
    /**
     * doc文件转换为html文件
     * @param sourceFileName doc文件名称
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws TransformerException
     */
    public void doc2Html(String sourceFileName) throws IOException, ParserConfigurationException, TransformerException {
        // 获取doc源文件目录
        String docPath = "G:/WordDB/";
        // 获取源文件名称
        String docName = sourceFileName.substring(0,sourceFileName.indexOf("."));
        // *.html 目标文件
        String targetFileName = "G:/Word2HtmlDB/"+docName+"/";
        File targetFile = new File(targetFileName);
        if(!targetFile.exists()){
            targetFile.mkdirs();
        }
        // html中图片存储
        String imagePathStr = "G:/Word2HtmlDB/"+docName+"/img/";
        File file = new File(imagePathStr);
        if(!file.exists()){
            file.mkdirs();
        }
        // 读取源文件
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(docPath+sourceFileName));
        // 创建文档对象
        Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        // 创建Word转换成Html的转换器对象
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        // 保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, name, width, height) -> {
            try(FileOutputStream out = new FileOutputStream(imagePathStr + name)){
                out.write(content);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return "img/" + name;
        });
        // 转换器对象处理Word对象
        wordToHtmlConverter.processDocument(wordDocument);
        // 获得html对象
        Document htmlDocument = wordToHtmlConverter.getDocument();
        // 根据html对象创建DOM对象
        DOMSource domSource = new DOMSource(htmlDocument);
        // 创建流对象结果
        StreamResult streamResult = new StreamResult(new File(targetFileName+docName+".html"));
        // 创建转换工厂
        TransformerFactory tf = TransformerFactory.newInstance();
        // 根据转换工厂创建转换器
        Transformer serializer = tf.newTransformer();
        // 转换器设置输出编码格式为UTF-8,转换方法
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        // 转换
        serializer.transform(domSource, streamResult);
    }

    /**
     * docx文件转换为html文件
     * @param sourceFileName docx文件名称
     * @throws IOException
     */
    public void docx2Html(String sourceFileName) throws IOException {
        // 获取docx源文件目录
        String docPath = "G:/WordDB/";
        // 获取源文件名称
        String docName = sourceFileName.substring(0,sourceFileName.indexOf("."));
        // *.html 目标文件
        String targetFileName = "G:/Word2HtmlDB/"+docName+"/";
        File targetFile = new File(targetFileName);
        if(!targetFile.exists()){
            targetFile.mkdirs();
        }
        // html中图片存储
        String imagePathStr = "G:/Word2HtmlDB/"+docName+"/img/";
        File file = new File(imagePathStr);
        if(!file.exists()){
            file.mkdirs();
        }
        // 定义输出流
        OutputStreamWriter outputStreamWriter = null;
        try {
            // 创建docx文档对象
            XWPFDocument document = new XWPFDocument(new FileInputStream(docPath+sourceFileName));
            // 创建XHTML操作对象
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePathStr)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver("img"));
            // 初始化输出流为*.html文件
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName+docName+".html"), "utf-8");
            // 创建xhtml文件转换器
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            // 转换
            xhtmlConverter.convert(document, outputStreamWriter, options);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outputStreamWriter != null) {
                outputStreamWriter.close();
            }
        }
    }

}
