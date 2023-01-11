package com.epal.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.WriteSheet;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

public class Excel2xmlUtils {

    /**
     * 单次缓存的数据量
     */
    public static final int BATCH_COUNT = 500;

    /**
     * excel转xml
     *
     * @param excelFilePath excel文件路径
     * @param xmlFileNames  需要生成的xml文件列表
     */
    public static void excelToXml(String excelFilePath, final String... xmlFileNames) {
        if (excelFilePath == null || excelFilePath.isEmpty()) {
            throw new IllegalArgumentException("filePath is null");
        }
        if (xmlFileNames == null || xmlFileNames.length == 0) {
            throw new IllegalArgumentException("xmlFileNames is null");
        }
        final File excelFile = new File(excelFilePath);
        if (!excelFile.exists()) {
            throw new IllegalArgumentException("the file " + excelFilePath + " don't exist");
        }
        EasyExcel.read(excelFilePath, new ReadListener<LinkedHashMap<Integer, String>>() {

            /**
             *临时存储
             */
            private List<LinkedHashMap<Integer, String>> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

            @Override
            public void invoke(LinkedHashMap<Integer, String> data, AnalysisContext context) {
                cachedDataList.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                /*for (int i = 0; i < cachedDataList.size(); i++) {
                    System.out.println(i + " -> " + cachedDataList.get(i));
                }*/
                generateXml(cachedDataList, xmlFileNames);
            }
        }).sheet(0).headRowNumber(0).doRead();
    }

    /**
     * xml转excel
     *
     * @param excelFilePath 需要生成的excel文件路径
     * @param xmlFileNames  源xml文件列表
     */
    public static void xml2Excel(String excelFilePath, String... xmlFileNames) {
        try {
            System.out.println("===xml文件转excel start===");
            final File excelFile = new File(excelFilePath);
            if (excelFile.exists()) {
                excelFile.delete();
            }
            final List<HashMap<String, String>> list = new ArrayList<>();
            final List<HashMap<String, HashMap<String, String>>> pluralsList = new ArrayList<>();
            for (String xmlFileName : xmlFileNames) {
                final DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                HashMap<String, String> valueMap = parseXml(factory.newDocumentBuilder(), xmlFileName);
                HashMap<String, HashMap<String, String>> valuesMap = parsePluralsXml(factory.newDocumentBuilder(), xmlFileName);
                list.add(valueMap);
                pluralsList.add(valuesMap);
            }

            final List<LinkedHashMap<Integer, String>> data = convertData(list);
            final List<LinkedHashMap<Integer, String>> pluralsData = convertPluralsData(pluralsList);

            try (ExcelWriter excelWriter = EasyExcel.write(excelFilePath).build()) {
                WriteSheet writeSheet = EasyExcel.writerSheet(0, "翻译").build();
                excelWriter.write(data, writeSheet);

                WriteSheet writePluralsSheet = EasyExcel.writerSheet(1, "复数翻译").build();
                excelWriter.write(pluralsData, writePluralsSheet);
            } catch (Exception e) {
                e.printStackTrace();
            }

            System.out.println("===end===");
        } catch (ParserConfigurationException e) {
            throw new RuntimeException(e);
        }
    }

    private static List<LinkedHashMap<Integer, String>> convertData(List<HashMap<String, String>> list) {
        List<LinkedHashMap<Integer, String>> result = new ArrayList<>();
        List<String> keyList = new ArrayList<>();
        for (HashMap<String, String> hashMap : list) {
            Set<String> strings = hashMap.keySet();
            for (String key : strings) {
                if (!keyList.contains(key)) {
                    keyList.add(key);
                }
            }
        }
        String value = "";
        for (String key : keyList) {
            LinkedHashMap<Integer, String> linkedHashMap = new LinkedHashMap<>();
            linkedHashMap.put(0, key);
            for (int i = 1; i < list.size() + 1; i++) {
                value = list.get(i - 1).get(key);
                if (value == null) {
                    value = "";
                }
                linkedHashMap.put(i, value);
            }
            result.add(linkedHashMap);
        }
        return result;
    }

    private static List<LinkedHashMap<Integer, String>> convertPluralsData(List<HashMap<String, HashMap<String, String>>> list) {
        List<LinkedHashMap<Integer, String>> result = new ArrayList<>();
        List<String> keyList = new ArrayList<>();
        for (HashMap<String, HashMap<String, String>> hashMap : list) {
            Set<String> strings = hashMap.keySet();
            for (String key : strings) {
                if (!keyList.contains(key)) {
                    keyList.add(key);
                }
            }
        }
        HashMap<String, String> value;
        for (String key : keyList) {
            for (int i = 1; i < list.size() + 1; i++) {
                value = list.get(i - 1).get(key);
                if (value != null) {
                    Set<String> strings = value.keySet();
                    for (String kv : strings) {
                        LinkedHashMap<Integer, String> linkedHashMap = new LinkedHashMap<>();
                        linkedHashMap.put(0, key);
                        linkedHashMap.put(1, kv);
                        linkedHashMap.put(2, value.get(kv));
                        result.add(linkedHashMap);
                    }
                }
            }
        }
        return result;
    }

    /**
     * 读取xml内容
     *
     * @param documentBuilder
     * @param xmlFileName
     * @return
     */
    private static HashMap<String, String> parseXml(DocumentBuilder documentBuilder, String xmlFileName) {
        try {
            final File f = new File(xmlFileName);
            final Document doc = documentBuilder.parse(f);
            final NodeList nl = doc.getElementsByTagName("string");
            Node item = null;
            HashMap<String, String> map = new LinkedHashMap<>();
            for (int i = 0; i < nl.getLength(); i++) {
                item = nl.item(i);
                if (item.getFirstChild() != null) {
                    map.put(item.getAttributes().getNamedItem("name").getNodeValue(), item.getFirstChild().getNodeValue());
                } else {
                    map.put(item.getAttributes().getNamedItem("name").getNodeValue(), null);
                }
            }
            return map;
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static HashMap<String, HashMap<String, String>> parsePluralsXml(DocumentBuilder documentBuilder, String xmlFileName) {
        try {
            final File f = new File(xmlFileName);
            final Document doc = documentBuilder.parse(f);
            final NodeList nl = doc.getElementsByTagName("plurals");
            Node item = null;
            HashMap<String, HashMap<String, String>> map = new LinkedHashMap<>();
            for (int i = 0; i < nl.getLength(); i++) {
                HashMap<String, String> keys = new LinkedHashMap<>();
                item = nl.item(i);
                NodeList childNodes = item.getChildNodes();
                for (int j = 0; j < childNodes.getLength(); j++) {
                    Node item1 = childNodes.item(j);
                    if (item1 instanceof Element) {
                        keys.put(item1.getAttributes().getNamedItem("quantity").getNodeValue(), item1.getFirstChild().getNodeValue());
                    }
                }
                map.put(item.getAttributes().getNamedItem("name").getNodeValue(), keys);
            }
            return map;
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /*
     <?xml version="1.0" encoding="utf-8"?>
     <resources>
        <string name="app_name">建图工具</string>
    </resources>
     */
    private static void generateXml(List<LinkedHashMap<Integer, String>> cachedDataList, String... xmlFileNames) {
        long startTime = System.currentTimeMillis();
        System.out.println("===start generate xml file===");

        List<StringBuilder> sbList = ListUtils.newArrayListWithExpectedSize(xmlFileNames.length);

        for (int i = 0; i < xmlFileNames.length; i++) {
            StringBuilder sb = new StringBuilder();
//            sb.append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
//            sb.append("\r\n");
//            sb.append("<resources>");
//            sb.append("\r\n");
            sbList.add(sb);
        }

        for (LinkedHashMap<Integer, String> linkedHashMap : cachedDataList) {
            for (int i1 = 1; i1 < linkedHashMap.entrySet().size(); i1++) {

                if (linkedHashMap.get(0) == null || linkedHashMap.get(0).isEmpty()) {
                    continue;
                }

                StringBuilder stringBuilder = sbList.get(i1 - 1);
                stringBuilder.append("\"");
                stringBuilder.append(linkedHashMap.get(0));
                stringBuilder.append("\" = \"");
                stringBuilder.append(linkedHashMap.get(i1));
                stringBuilder.append("\";");
                stringBuilder.append("\r\n");
            }
        }

//        for (StringBuilder sb : sbList) {
//            sb.append("</resources>");
//        }
        for (int i = 0; i < sbList.size(); i++) {
            writeString2File(xmlFileNames[i], sbList.get(i).toString());
        }
        System.out.println("generateXml cost time is " + (System.currentTimeMillis() - startTime));
        System.out.println("===end===");
    }

    private static void writeString2File(String filePath, String text) {
        final File file = new File(filePath);
        FileOutputStream fos = null;
        try {
            if (!file.exists()) {
                boolean createNewFile = file.createNewFile();
                if (!createNewFile) {
                    System.out.println("create file error, file path is " + filePath);
                    return;
                }
            }
            fos = new FileOutputStream(file);
            fos.write(text.getBytes(StandardCharsets.UTF_8));
            fos.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 根据英文翻译，合并翻译excel文档
     *
     * @param sourceFilePath    应用excel文档
     * @param translateFilePath 多语言翻译文档
     * @param mergeFilePath     输出文档
     */
    public static void compareMergeFile(String sourceFilePath, String translateFilePath, String mergeFilePath) {
        System.out.println("===start compare merge excel file===");
        if (sourceFilePath == null || sourceFilePath.isEmpty()) {
            throw new IllegalArgumentException("sourceFilePath is null");
        }
        final File excelFile = new File(sourceFilePath);
        if (!excelFile.exists()) {
            throw new IllegalArgumentException("the sourceFilePath " + sourceFilePath + " don't exist");
        }

        if (translateFilePath == null || translateFilePath.isEmpty()) {
            throw new IllegalArgumentException("translateFilePath is null");
        }
        final File translateExcelFile = new File(translateFilePath);
        if (!translateExcelFile.exists()) {
            throw new IllegalArgumentException("the translateFilePath " + translateExcelFile + " don't exist");
        }

        final File outFile = new File(mergeFilePath);
        if (outFile.exists()) {
            outFile.delete();
        }

        //临时存储（项目原英文翻译 0-key 1-英文）
        final List<LinkedHashMap<Integer, String>> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
        EasyExcel.read(sourceFilePath, new ReadListener<LinkedHashMap<Integer, String>>() {

            @Override
            public void invoke(LinkedHashMap<Integer, String> data, AnalysisContext context) {
                cachedDataList.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                /*for (int i = 0; i < cachedDataList.size(); i++) {
                    System.out.println(i + " -> " + cachedDataList.get(i));
                }*/
                System.out.println("读取完毕->项目原英文文档");
            }
        }).sheet(0).headRowNumber(0).doRead();

        //临时存储（多语言翻译文档 0-英文 1-西语）
        final List<LinkedHashMap<Integer, String>> translateDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
        EasyExcel.read(translateFilePath, new ReadListener<LinkedHashMap<Integer, String>>() {

            @Override
            public void invoke(LinkedHashMap<Integer, String> data, AnalysisContext context) {
                translateDataList.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                /*for (int i = 0; i < translateDataList.size(); i++) {
                    System.out.println(i + " -> " + translateDataList.get(i));
                }*/
                System.out.println("读取完毕->多语言翻译文档");
            }
        }).sheet(0).headRowNumber(0).doRead();

        final List<LinkedHashMap<Integer, String>> noneDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
        //开始比对英文，自动填充西语
        System.out.println("开始合并文件");
        for (LinkedHashMap<Integer, String> hashMap : cachedDataList) {
            String en = hashMap.get(1);
            if (en != null) {
                for (LinkedHashMap<Integer, String> translateHashMap : translateDataList) {
                    String ten = translateHashMap.get(0);
                    if (ten != null) {
                        if (compareString(en, ten)) {//在翻译文档中找到对应的英文了
                            String txy = translateHashMap.get(1);
                            if (txy != null) {
                                hashMap.put(2, txy);
                                break;//退出此循环
                            }
                        } else {
                            //手动处理一些翻译
                            String txy = translateHashMap.get(1);
                            String compareText = compareSpacial(en, ten, txy);
                            if (compareText != null) {
                                hashMap.put(2, compareText);
                                break;//退出此循环
                            }
                        }
                    }
                }

                //没有比对到
                if (hashMap.get(2) == null) {
                    noneDataList.add(hashMap);
                }

                //没有在翻译文档中找到匹配的英文，补齐长度
                if (hashMap.size() == 2) {
                    hashMap.put(2, null);
                }
            }
        }

        try (ExcelWriter excelWriter = EasyExcel.write(mergeFilePath).build()) {
            WriteSheet writeSheet = EasyExcel.writerSheet(0, "比对翻译").build();
            excelWriter.write(cachedDataList, writeSheet);

            WriteSheet writePluralsSheet = EasyExcel.writerSheet(1, "缺失翻译").build();
            excelWriter.write(noneDataList, writePluralsSheet);
        } catch (Exception e) {
            e.printStackTrace();
        }
        /*EasyExcel.write(mergeFilePath)
                .sheet("比对翻译")
                .doWrite(cachedDataList);*/
        System.out.println("===end===");
    }

    private static boolean compareString(String s1, String s2) {
        if (s1.equalsIgnoreCase(s2)) {
            return true;
        }

        s1 = convert(s1);
        s2 = convert(s2);

//        if (s1.contains("can't find that user") && s2.contains("can't find that user")) {
//            System.err.println(s1);
//            System.err.println(s2);
//        }

        return s1.equalsIgnoreCase(s2);
    }

    private static String compareSpacial(String s1, String s2, String s3) {
        if (s1.equals("×%s")) {
            return "×%s";
        }

        if (s3 == null) {
            return null;
        }

        s1 = convert(s1);
        s2 = convert(s2);
        String rs1;
        String rs2;
        if (s2.contains("%@")) {
            if (s1.contains("%s")) {
                rs2 = s2.replaceAll("%@", "%s");
            } else {
                rs2 = s2.replaceAll("%@", "%d");
            }
            if (s1.equals(rs2)) {
                if (s1.contains("%s")) {
                    return s3.replaceAll("%@", "%s");
                } else {
                    return s3.replaceAll("%@", "%d");
                }
            }
        }

        if (s1.startsWith("Oops...") && s2.startsWith("Oops...")) {
            rs1 = s1.replace("Oops...", "").trim();
            rs2 = s2.replace("Oops...", "").trim();

            if (rs1.contains("can't find that user") && rs2.contains("can't find that user")) {
                System.err.println(rs1);
                System.err.println(rs2);
            }
            if (rs1.equals(rs2)) {
                return s3;
            }
        }

        if (s1.startsWith("*")) {
            rs1 = s1.replace("*", "").trim();
            if (rs1.equals(s2)) {
                return "* " + s3;
            }
        }

        if (s1.startsWith("%d")) {
            rs1 = s1.replace("%d", "").trim();
            if (rs1.equals(s2)) {
                return "%d " + s3;
            }
        }

        if (s1.startsWith("%s")) {
            rs1 = s1.replace("%s", "").trim();
            if (rs1.equals(s2)) {
                return "%s " + s3;
            }
        }

        if (s1.endsWith("%d")) {
            rs1 = s1.replace("%d", "").trim();
            if (rs1.equals(s2)) {
                return s3 + " %d";
            }
        }

        if (s1.endsWith("%s")) {
            rs1 = s1.replace("%s", "").trim();
            if (rs1.equals(s2)) {
                return s3 + " %s";
            }
        }

        if (s1.endsWith(": %s")) {
            rs1 = s1.replace(": %s", "").trim();
            if (rs1.equals(s2)) {
                return s3 + ": %s";
            }
        }

        if (s1.endsWith("(%d)")) {
            rs1 = s1.replace("(%d)", "").trim();
            if (rs1.equals(s2)) {
                return s3 + "(%d)";
            }
        }

        if (s1.endsWith("(%s)")) {
            rs1 = s1.replace("(%s)", "").trim();
            if (rs1.equals(s2)) {
                return s3 + "(%s)";
            }
        }

        if (s1.endsWith(":")) {
            rs1 = s1.replace(":", "").trim();
            if (rs1.equals(s2)) {
                return s3 + ":";
            }
        }

        if (s1.endsWith("…")) {
            rs1 = s1.replace("…", "").trim();
            if (rs1.equals(s2)) {
                return s3 + "…";
            }
        }

        if (s1.endsWith("...")) {
            rs1 = s1.replace("...", "").trim();
            if (rs1.equals(s2)) {
                return s3 + "...";
            }
        }

        if (s1.endsWith(">")) {
            rs1 = s1.replace(">", "").trim();
            if (rs1.equals(s2)) {
                return s3 + " >";
            }
        }

        if (s1.endsWith(">") && s2.endsWith(">")) {
            rs1 = s1.replace(">", "").trim();
            rs2 = s2.replace(">", "").trim();
            if (rs1.equals(rs2)) {
                return s3;
            }
        }

        if (s1.endsWith(" ")) {
            rs1 = s1.substring(0, s1.length() - 1);
            if (rs1.equals(s2)) {
                return s3 + " ";
            }
        }

        if (s1.equals("%s · %s Reviews") && s2.equals("Reviews")) {
            return "%s · %s " + s3;
        }

        if (s1.equals("2.00 to Draw") && s2.equals("%@ to Draw")) {
            return s3.replace("%@", "2.00");
        }

        if (s1.equals("+%s exp") && s2.equals("exp")) {
            return "+%s " + s3;
        }

        if (s1.equals("Awards · %d >") && s2.equals("Awards")) {
            return s3 + " · %d >";
        }
        if (s1.equals("Active Days: %s") && s2.equals("Active Days:{{val}}")) {
            return s3.replace("{{val}}", "%s");
        }
        if (s1.equals("Performance\\n%.1f") && s2.equals("Performance")) {
            return s3 + "\\n%.1f";
        }
        if (s1.equals("Friendliness\\n%.1f") && s2.equals("Friendliness")) {
            return s3 + "\\n%.1f";
        }
        if (s1.equals("Responsive\\n%.1f") && s2.equals("Responsive")) {
            return s3 + "\\n%.1f";
        }
        if (s1.equals("Enjoyment\\n%.1f") && s2.equals("Enjoyment")) {
            return s3 + "\\n%.1f";
        }
        if (s1.equals("(Optional)") && s2.equals("Optional")) {
            return "(" + s3 + ")";
        }
        if (s1.equals(" (Optional)") && s2.equals("Optional")) {
            return " (" + s3 + ")";
        }
        if (s1.equals("(Buff)") && s2.equals("Buff")) {
            return "(" + s3 + ")";
        }
        if (s1.equals("%s%% Off") && s2.equalsIgnoreCase("Off")) {
            return "%s%% " + s3;
        }
        if (s1.equals("+%s%% For Newcomer") && s2.equals("For Newcomer")) {
            return "%s%% " + s3;
        }
        if (s1.equals("Service Types · %d") && s2.equals("Service Types")) {
            return s3 + " · %d";
        }
        if (s1.equals("Resend (%s s)") && s2.equals("Resend")) {
            return s3 + " (%s s)";
        }
        if (s1.equals("Ordered \\\"%s\\\" %s") && s2.equals("Ordered")) {
            return s3 + " \\\"%s\\\" %s";
        }
        if (s1.equals("\\n%s - Voice Rating: %d") && s2.equals("Voice Rating")) {
            return "\\n%s - " + s3 + ": %d";
        }
        if (s1.equals("\\n%s - Cover Rating: %d") && s2.equals("Cover Rating")) {
            return "\\n%s - " + s3 + ": %d";
        }
        if (s1.endsWith("18 years old or above") && s2.endsWith("18 years old or above")) {
            return s3;
        }
        if (s1.endsWith("Able to provide service in at least 1 category on E-Pal") && s2.endsWith("Able to provide service in at least 1 category on E-Pal")) {
            return s3;
        }
        if (s1.endsWith("Enthusiastic, good at communication, non-toxic and do not conduct NSFW services") && s2.endsWith("Enthusiastic, good at communication, non-toxic and do not conduct NSFW services")) {
            return s3;
        }
        if (s1.endsWith("$%s/Month") && s2.equalsIgnoreCase("Month")) {
            return "$%s/" + s3;
        }
        if (s1.equals("15Min") && s2.equalsIgnoreCase("Min")) {
            return "15" + s3;
        }
        if (s1.equals("30Min") && s2.equalsIgnoreCase("Min")) {
            return "30" + s3;
        }
        if (s1.equals("Influencer Project ") && s2.equals("Influencer Project")) {
            return s3 + " ";
        }
        if (s1.startsWith("After completing this week's entrance effect task, the Entrance Effect will be sent on Monday 00:00 PST") && s2.startsWith("After completing this week's entrance effect task, the Entrance Effect will be sent on Monday 00:00 PST")) {
            return s3 + "%s/%s";
        }
        return null;
    }

    private static String convert(String text) {
        text = text.trim();
        if (text.endsWith(".") || text.endsWith("?") || text.endsWith("？")) {
            if (!text.endsWith("...")) {
                text = text.substring(0, text.length() - 1);
            }
        }
        if (text.contains("\\'")) {
            text = text.replace("\\'", "'");
        }
        if (text.contains("\\\"")) {
            text = text.replace("\\\"", "\"");
        }
        if (text.contains("’")) {
            text = text.replaceAll("’", "'");
        }
        if (text.contains("…")) {
            text = text.replaceAll("…", "...");
        }
        if (text.contains("“")) {
            text = text.replaceAll("“", "\"");
        }
        if (text.contains("”")) {
            text = text.replaceAll("”", "\"");
        }
        if (text.contains("\r")) {
            text = text.replaceAll("\r", "");
        }
        return text.trim();
    }
}