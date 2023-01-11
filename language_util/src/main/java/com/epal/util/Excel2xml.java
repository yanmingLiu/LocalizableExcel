package com.epal.util;

import java.io.File;

public class Excel2xml {
    public static void main(String[] args) {
        //找到项目资源文件目录
        String rootPath = System.getProperty("user.dir");
        String dirFile = rootPath.concat("/language");
        File file = new File(dirFile);
        if (!file.exists()) {
            file.mkdirs();
        }
        //xml转excel（读取项目中的string.xml，生成对应的excel文件）
        String sourceFilePath = dirFile.concat("/source.xlsx");
        String sourceXmlPath = rootPath.concat("/app/src/main/res/values/strings.xml");
        String sourceXmlPath2 = rootPath.concat("/app/src/main/res/values-es/strings.xml");
        String sourceXmlPath3 = rootPath.concat("/app/src/main/res/values-de/strings.xml");
        String sourceXmlPath4 = rootPath.concat("/app/src/main/res/values-fr/strings.xml");//
        String sourceXmlPath5 = rootPath.concat("/app/src/main/res/values-tr/strings.xml");//
        String sourceXmlPath6 = rootPath.concat("/app/src/main/res/values-it/strings.xml");//
        //Excel2xmlUtils.xml2Excel(sourceFilePath, sourceXmlPath, sourceXmlPath2, sourceXmlPath3, sourceXmlPath4, sourceXmlPath5, sourceXmlPath6);

        //比对翻译文案（比对strings生成的excel和翻译文档excel） TODO sourceFilePath（0-key 1-比对语言）
        String translateFilePath = dirFile.concat("/translate.xlsx");//TODO 多语言翻译文档（ 0-比对语言 1-翻译语言）
        //String mergeFilePath = dirFile.concat("/merge.xlsx");//合并文件
        //Excel2xmlUtils.compareMergeFile(sourceFilePath, translateFilePath, mergeFilePath);

        //将翻译文件合并的excel转换为多语言的xml
        String exportFilePath1 = dirFile.concat("/strings-en.xml");//语言1
        String exportFilePath2 = dirFile.concat("/strings-es.xml");//语言2
        String exportFilePath3 = dirFile.concat("/strings-de.xml");//语言2
        String exportFilePath4 = dirFile.concat("/strings-fr.xml");//语言2
        String exportFilePath5 = dirFile.concat("/strings-tr.xml");//语言2
        String exportFilePath6 = dirFile.concat("/strings-it.xml");//语言2
        String mergeFilePath = dirFile.concat("/translate.xlsx");//合并文件
        Excel2xmlUtils.excelToXml(mergeFilePath, exportFilePath1, exportFilePath2, exportFilePath3, exportFilePath4, exportFilePath5, exportFilePath6);
    }
}