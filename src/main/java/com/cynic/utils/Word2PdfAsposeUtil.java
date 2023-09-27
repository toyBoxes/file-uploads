package com.cynic.utils;


import cn.hutool.core.io.resource.ClassPathResource;
import cn.hutool.core.io.resource.Resource;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

public class Word2PdfAsposeUtil {
    // 添加日志
    private static final Logger logger = LoggerFactory.getLogger(Word2PdfAsposeUtil.class);
    // 读取license.xml的内容
    public static boolean getLicense() {
        boolean result = false;
        Resource resource = new ClassPathResource("static/license.xml");
        try (InputStream is = resource.getStream()) {
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
    /**
     *
     * @param inPath 源文件路径 .docx文件
     * @param outPath 生成的文件 .pdf文件
     */
    public static boolean doc2pdf(String inPath, String outPath) {
        if (!getLicense()) {
            logger.error("license获取失败");
            return false;
        }
        try (FileOutputStream os = new FileOutputStream(new File(outPath))) {
            long old = System.currentTimeMillis();
            Document doc = new Document(inPath);
            doc.save(os, SaveFormat.PDF);
            long now = System.currentTimeMillis();
            logger.info("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒");
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }
    // main 方法测试工具类
    public static void main(String[] args) {
        //调用授权方法 源文件路径
        String sourceFile = "C:\\Users\\12052320\\Documents\\个人网银文档\\YF2020013_金融互联网平台升级项目_业务需求说明书_个人网银渠道__钱莞家基金分册V0.24.docx";
        // 生成的路径
        String targetFile = "C:\\Users\\12052320\\Documents\\个人网银文档\\YF2020013_金融互联网平台升级项目_业务需求说明书_个人网银渠道__钱莞家基金分册V0.24.pdf";
        // 调用方法
        doc2pdf(sourceFile,targetFile);
    }
}

