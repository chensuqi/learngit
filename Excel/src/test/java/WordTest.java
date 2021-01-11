import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class WordTest {
    public static void main(String[] args) throws Exception {
        StringBuilder sb = new StringBuilder();
        sb.append("<h1>标题1</h1>")
                .append("<h2>标题2</h2>")
                .append("<table style=\"border-collapse: collapse; width: 549.4pt; border: none; text-align: justify; font-family: 'Times New Roman'; font-size: 10pt;\" border=\"1\" cellspacing=\"0\">")
                .append("<tbody>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 549.4000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; background: #6d0001; border: 1.0000pt solid windowtext;\" colspan=\"2\" valign=\"top\" width=\"549\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">计划</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">本周完成情况</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: 1.0000pt solid windowtext; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">实际完成</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">0%</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">：2020-12-28</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;</span><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">：2021-01-01</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">工作内容</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 12.0000pt;\">1111111111111111111111</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">进展反馈</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">实际完成</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">0%</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">：2020-12-28</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;</span><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">：2021-01-01</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">工作内容</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 12.0000pt;\">1111111111111111111111</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">进展反馈</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">下周工作计划</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">预计完成</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">0%</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2021-01-04</span><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">&nbsp;&nbsp;</span><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2021-01-08</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">工作内容</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">1111111111111111111111</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">备注</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">预计完成</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></strong><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;</span><span style=\"font-family: 宋体; font-size: 12.0000pt;\">0%</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2021-01-04</span><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">&nbsp;&nbsp;</span><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2021-01-08</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">工作内容</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">1111111111111111111111</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">备注</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">出差计划</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2020-12-28</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2020-12-31</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">出差地点</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">厦门</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">交通方式</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">火车</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">具体安排</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">钱钱钱</span></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">&nbsp;</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">开始日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2020-12-28</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">结束日期：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">2020-12-31</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">出差地点</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">厦门</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">交通方式</span></strong><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">：</span></strong><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">火车</span></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">具体安排</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">钱钱钱</span></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 14.4500pt;\">")
                .append("<td style=\"width: 130.3500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"130\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">工作总结</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 419.0500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"419\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-weight: normal; font-size: 12.0000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr style=\"height: 4.5000pt;\">")
                .append("<td style=\"width: 549.4000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext; background: #6d0001;\" colspan=\"2\" valign=\"top\" width=\"549\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">市场信息</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr>")
                .append("<td style=\"width: 549.4000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" colspan=\"2\" valign=\"top\" width=\"549\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 12.0000pt;\">价格回顾</span></strong></p>")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 12.0000pt;\">&nbsp;</span></p>")
                .append("<table style=\"border-collapse: collapse; border: none; text-align: justify; font-family: 'Times New Roman'; font-size: 10pt;\" border=\"1\" cellspacing=\"0\">")
                .append("<tbody>")
                .append("<tr>")
                .append("<td style=\"width: 70.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; background: #dce6f2; border: 1.0000pt solid windowtext;\" valign=\"top\" width=\"70\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 7.5000pt;\">产品</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 85.1000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: 1.0000pt solid windowtext; border-bottom: 1.0000pt solid windowtext; background: #dce6f2;\" valign=\"top\" width=\"85\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 7.5000pt;\">周五报价</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 56.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: 1.0000pt solid windowtext; border-bottom: 1.0000pt solid windowtext; background: #dce6f2;\" valign=\"top\" width=\"56\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 7.5000pt;\">币种</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: 1.0000pt solid windowtext; border-bottom: 1.0000pt solid windowtext; background: #dce6f2;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 7.5000pt;\">省份</span></strong></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: 1.0000pt solid windowtext; border-bottom: 1.0000pt solid windowtext; background: #dce6f2;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: center; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\" align=\"center\"><strong><span style=\"font-family: 宋体; font-weight: bold; font-size: 7.5000pt;\">备注</span></strong></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr>")
                .append("<td style=\"width: 70.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"70\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">农副产品</span></p>")
                .append("</td>")
                .append("<td style=\"width: 85.1000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"85\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 56.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"56\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("</tr>")
                .append("<tr>")
                .append("<td style=\"width: 70.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: 1.0000pt solid windowtext; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"70\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">玉米油</span></p>")
                .append("</td>")
                .append("<td style=\"width: 85.1000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"85\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 56.9500pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"56\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("<td style=\"width: 71.0000pt; padding: 0.0000pt 5.4000pt 0.0000pt 5.4000pt; border-left: none; border-right: 1.0000pt solid windowtext; border-top: none; border-bottom: 1.0000pt solid windowtext;\" valign=\"top\" width=\"71\">")
                .append("<p style=\"text-align: left; margin: 0pt 0pt 0.0001pt; font-family: 'Times New Roman'; font-size: 12pt;\"><span style=\"font-family: 宋体; font-size: 7.5000pt;\">&nbsp;</span></p>")
                .append("</td>")
                .append("</tr>")
                .append("</tbody>")
                .append("</table>")
        ;
        String content = "<html><head></head><body>" + sb.toString() + "</body></html>";
        InputStream is = new ByteArrayInputStream(content.getBytes("GBK"));
        String FileName = "D:\\fuyy\\Desktop\\测试1.docx";
        OutputStream os = new FileOutputStream(FileName);
//        inputStreamToWord(is, os);


        XWPFDocument xwpfDocument = new XWPFDocument();

        XWPFParagraph xwpfParagraph = xwpfDocument.createParagraph();

        XWPFRun xwpfRun = xwpfParagraph.createRun();

        xwpfRun.setText(content);



        xwpfDocument.write(os);

        os.close();
        xwpfDocument.close();





    }

    /**
     * 把is写入到对应的word输出流os中
     * 不考虑异常的捕获，直接抛出
     */
    private static void inputStreamToWord(InputStream is, OutputStream os) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem();
        //对应于org.apache.poi.hdf.extractor.WordDocument
        fs.createDocument(is, "WordDocument");
        fs.writeFilesystem(os);
        os.close();
        is.close();
    }

    /**
     * 把输入流里面的内容以UTF-8编码当文本取出。
     * 不考虑异常，直接抛出
     */
    private static String getContent(InputStream... ises) throws IOException {
        if (ises != null) {
            StringBuilder result = new StringBuilder();
            BufferedReader br;
            String line;
            for (InputStream is : ises) {
                br = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8));
                while ((line = br.readLine()) != null) {
                    result.append(line);
                }
            }
            return result.toString();
        }
        return null;
    }
}
