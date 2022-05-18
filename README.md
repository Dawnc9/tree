package com.huawei.allinone.activitiForm.majorProjects.controller;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;

import org.junit.Test;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/word")
public class WordTest {
    @Test
    public void test() throws IOException, XDocReportException {
        generateWord();
    }
    public void generateWord() throws IOException, XDocReportException {
        //获取Word模板，模板存放路径在项目的resources目录下
        InputStream ins = new FileInputStream("D:\\评审纪要模板.docx");
        //注册xdocreport实例并加载FreeMarker模板引擎
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(ins,
            TemplateEngineKind.Freemarker);
        //创建xdocreport上下文对象
        IContext context = report.createContext();
        //创建要替换的文本变量
        context.put("createProject_problemOverview", "长安银行AIX主机异构/LVM迁移方案");
        Map<String,Object> good = new HashMap();
        good.put("num","1321");
        good.put("type","大幅度范德萨发生");
        List<Map<String,Object>> goodsList = new ArrayList<>();
        Map<String,Object> goods1 = new HashMap();
        goods1.put("num",1);
        goods1.put("type","臭美毁肤");
        Map<String,Object> goods2 = new HashMap<>();
        goods2.put("num",2);
        goods2.put("type","女装");
        Map<String,Object> goods3 = new HashMap();
        goods3.put("num",3);
        goods3.put("type","手机");
        goodsList.add(goods1);
        goodsList.add(goods2);
        goodsList.add(goods3);
        context.put("goods", goodsList);
        context.put("good", good);
        //创建字段元数据
        FieldsMetadata fm = report.createFieldsMetadata();
        //Word模板中的表格数据对应的集合类型
        fm.load("good", HashMap.class, false);
        fm.load("goods", HashMap.class, true);
        //输出到本地目录
        FileOutputStream out = new FileOutputStream(new File("D:\\评审纪要.docx"));
        report.process(context, out);
    }
}
