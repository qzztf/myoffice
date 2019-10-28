package cn.sexycode.office.util;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(WordUtil.class);

    public void replaceText(XWPFDocument doc, Map<String, Object> params) {
        /*Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }*/

        processParagraphs(doc.getParagraphs(), params);
        replaceInTable(doc, params);
    }

    private void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        // 处理表格
        Iterator<XWPFTable> it = doc.getTablesIterator();
        while (it.hasNext()) {
            XWPFTable table = it.next();
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
                    processParagraphs(paragraphListTable, params);
                }
            }
        }
    }

    /**
     * 替换占位符
     *
     * @param text
     * @param map
     * @return
     */
    private String replaceText(String text, Map<String, Object> map) {
        if (text != null) {
            Matcher matcher = matcher(text);
            if (matcher.find()) {
                while ((matcher = matcher(text)).find()) {
                    text = matcher.replaceFirst(String.valueOf(map.get(matcher.group(1))));
                }
                if (text.equals("null")) {
                    text = "";
                }
            }
        }
        return text;
    }

    private void processParagraphs(List<XWPFParagraph> paragraphs, Map<String, Object> params) {
        for (XWPFParagraph paragraph : paragraphs) {
            replaceInPara(paragraph, params);
        }
    }

    public static void main(String[] args) throws Exception {
        WordUtil util = new WordUtil();
        XWPFDocument doc = new XWPFDocument(
                new FileInputStream("d:\\x5\\房屋抵押履约保证保险投保单（紫金-提放保）20190202.docx_1551061403282.docx"));
        Map<String, Object> map = new HashMap<>();
        map.put("bankName", "111111");
        map.put("bankAddress", "223434242");
        map.put("creditCode", "223434242bbbbb");
        util.replaceText(doc, map);
        doc.write(new FileOutputStream("d:/x5/demo2222333.docx"));
    }

    private void replaceInPara(XWPFParagraph para, Map<String, Object> param) {
        List<XWPFRun> runs;
        String tempString = "";
        char lastChar = ' ';
        if (matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            Set<XWPFRun> runSet = new HashSet<>();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                LOGGER.debug("=======>" + text);
                if (text == null)
                    continue;
                String replacedText = replaceText(text, param);
                if (!text.equalsIgnoreCase(replacedText)) {
                    //已经替换成功
                    run.setText("", 0);
                    setWrap(replacedText, para, run);
                    continue;
                }
                text = replacedText;
                run.setText(text, 0);

                for (int i = 0; i < text.length(); i++) {
                    char ch = text.charAt(i);
                    if (ch == '$') {
                        runSet = new HashSet<>();
                        runSet.add(run);
                        tempString = text;
                    } else if (ch == '{') {
                        if (lastChar == '$') {
                            if (!runSet.contains(run)) {
                                runSet.add(run);
                                tempString = tempString + text;
                            }
                        } else {
                            runSet = new HashSet<>();
                            tempString = "";
                        }
                    } else if (ch == '}') {
                        if (tempString.contains("${")) {
                            if (!runSet.contains(run)) {
                                runSet.add(run);
                                tempString = tempString + text;
                            }
                        } else {
                            runSet = new HashSet<>();
                            tempString = "";
                        }
                        if (runSet.size() > 0) {
                            String replaceText = replaceText(tempString, param);
                            if (!replaceText.equals(tempString)) {
                                int index = 0;
                                XWPFRun aRun = null;
                                for (XWPFRun tempRun : runSet) {
                                    tempRun.setText("", 0);
                                    if (index == 0) {
                                        aRun = tempRun;
                                    }
                                    index++;
                                }
                                if (aRun != null) {
                                    setWrap(replaceText, para, aRun);
                                }
                            }
                            runSet = new HashSet<>();
                            tempString = "";
                        }
                    } else {
                        if (runSet.size() <= 0)
                            continue;
                        if (runSet.contains(run))
                            continue;
                        runSet.add(run);
                        tempString = tempString + text;
                    }
                    lastChar = ch;
                }
            }
        }

    }

    private void setWrap(Object value, XWPFParagraph paragraph, XWPFRun run) {
        if (((String) value).indexOf("\r") > 0) {
            //设置换行
            String[] text = value.toString().split("\r");
            //            paragraph.removeRun(0);
            //            run = paragraph.insertNewRun(0);
            for (int f = 0; f < text.length; f++) {
                if (f == 0) {
                    //此处不缩进因为word模板已经缩进了。
                    run.setText(text[f].trim());
                } else {
                    run.addCarriageReturn();//硬回车
                    //注意：wps换行首行缩进是三个空格符，office要的话可以用 run.addTab();缩进或者四个空格符
                    run.setText(text[f]);
                }
            }
        } else {
            run.setText((String) value);
        }
    }

    private void setStyle(XWPFDocument doc, XWPFDocument tempdoc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        Iterator<XWPFParagraph> iterator2 = tempdoc.getParagraphsIterator();
        XWPFParagraph para;
        XWPFParagraph para2;
        while (iterator.hasNext()) {
            para = iterator.next();
            para2 = iterator2.next();
            this.styleInPara(para, para2, params);
        }
    }

    private void styleInPara(XWPFParagraph para, XWPFParagraph para2, Map<String, Object> params) {
        List<XWPFRun> runs;
        List<XWPFRun> runs2;
        Matcher matcher;
        if (matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            runs2 = para2.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                XWPFRun run2 = runs2.get(i);
                String runText = run.toString();
                matcher = matcher(runText);
                if (matcher.find()) {
                    //按模板文件格式设置格式
                    run2.getCTR().setRPr(run.getCTR().getRPr());
                }
            }
        }
    }

    private Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private Matcher matcher(String str) {
        return pattern.matcher(str);
    }

    /**
     * 关闭输入流
     *
     * @param is
     */
    public void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    public void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
