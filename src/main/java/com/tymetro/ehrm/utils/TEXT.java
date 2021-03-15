/*****************************************************************************************
 Class Name: TEXT
 Package: com.kjsoft.common
 Purpose: Text Utility
 Memo:
 History:
 Mod #	Date		By			Description
 -------	----------	--------	--------------------------------------------------
 *000    2019-05-07 	Jimmy Lee	Initial Code

 Copyright (c) 2019, KJSOFT. All rights reserved.
 *****************************************************************************************/

package com.tymetro.ehrm.utils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.StringReader;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.http.HttpServletRequest;

/**
 * 處理字串，文字轉換，HTML等相關功能
 * @author Jimmy Lee ( <a href="mailto:jimmy@kjsoft.com.tw">jimmy@kjsoft.com.tw</a> )
 * @version 1.0
 */
public class TEXT {

    /**
     * 檢查輸入字串是否合法, 若為NULL或是空白則傳回true
     * @param 檢核字串
     * @return NULL或是空白字串則傳回true
     */
    public static boolean isEmpty(String sInput) {
        if (sInput == null || sInput.length() == 0)
            return true;
        return false;
    }

    /**
     * 檢查輸入字串陣列是否合法, 若為NULL或是空白則傳回true
     * @param 檢核字串
     * @return NULL或是空白字串則傳回true
     */
    public static boolean isEmpty(String[] sInput) {
        if (sInput == null || sInput.length == 0)
            return true;
        return false;
    }


    /**
     * 檢查輸入字串是否為NULL
     * @param 檢核字串
     * @return NULL則傳回true
     */
    public static boolean isNull(String sInput) {
        if (sInput == null) {
            return true;
        }
        return false;
    }

    /**
     * 檢查是否為NULL, 若為NULL則傳回""空字串
     * @param 檢核字串
     * @return 回值字串
     */
    public static String checkNull(String sInput) {
        if (sInput == null)
            return "";
        return sInput;
    }

    /**
     * 取得網頁參數
     * @param Http Request物件
     * @param 網頁參數Name值
     * @return 網頁參數Value值或是空字串
     */
    public static String getParameter(HttpServletRequest oRequest, String sParName) {
        String parameter = oRequest.getParameter(sParName);
        return checkNull(parameter);
    }

    /**
     * 比較兩個字串是否完全相等，若兩者皆為Null也是相等
     * @param 來源字串
     * @param 目的字串
     * @return true/false
     */
    public static boolean compare(String strA, String strB) {
        if (strA == null && strB == null) {
            return true;
        } else if (strA == null){
            return strB.equals(strA);
        } else {
            return strA.equals(strB);
        }
    }

    /**
     * 轉換BIG5字串為ISO字串
     * @param BIG5字串
     * @return ISO字串
     */
    public static String convertBIG5toISO(String sInput) {
        try {
            if (sInput == null)
                return sInput;
            return new String(sInput.getBytes("BIG5"), "ISO-8859-1");
        } catch (Exception e) {
            System.out.println("BIG5轉ISO字串轉換錯誤:" + e.toString());
            return sInput;
        }
    }

    /**
     * 轉換ISO字串為BIG5字串
     * @param ISO字串
     * @return BIG5字串
     */
    public static String convertISOtoBIG5(String sInput) {
        try {
            if (sInput == null)
                return sInput;
            return new String(sInput.getBytes("ISO-8859-1"), "BIG5");
        } catch (Exception e) {
            System.out.println("ISO轉BIG5字串轉換錯誤:" + e.toString());
            return sInput;
        }
    }

    /**
     * 轉換ISO字串為UTF8字串
     * @param ISO字串
     * @return UTF8字串
     */
    public static String convertISOtoUTF8(String sInput) {
        try {
            if (sInput == null)
                return sInput;
            return new String(sInput.getBytes("ISO-8859-1"), "UTF-8");
        } catch (Exception e) {
            System.out.println("ISO轉UTF8字串轉換錯誤:" + e.toString());
            return sInput;
        }
    }

    /**
     * 轉換BIG5字串為UTF8字串
     * @param BIG5字串
     * @return UTF8字串
     */
    public static String convertBIG5toUTF8(String sInput) {
        try {
            if (sInput == null)
                return sInput;
            return new String(sInput.getBytes("BIG5"), "UTF-8");
        } catch (Exception e) {
            System.out.println("BIG5轉UTF8字串轉換錯誤:" + e.toString());
            return sInput;
        }
    }

    /**
     * 轉換整數變成文字
     * @param 整數
     * @return 文字
     */
    public static String toString(int iInteger) {
        return Integer.toString(iInteger);
    }

    /**
     * 轉換文字成為整數
     * @param 文字
     * @return 整數
     */
    public static int toInt(String sInteger) {
        try {
            if (isEmpty(sInteger)) {
                return 0;
            } else {
                return Integer.parseInt(sInteger);
            }
        } catch (Exception e) {
            return 0;
        }
    }

    /**
     * 轉換文字成為浮點數
     * @param 文字
     * @return 浮點數
     */
    public static float toFloat(String sFloat) {
        if (isEmpty(sFloat)) {
            return 0;
        } else {
            return Float.parseFloat(sFloat);
        }
    }

    /**
     * 轉換文字成為布林值
     * @param 文字
     * @return 布林值
     */
    public static boolean toBoolean(String sInteger) {
        if ("true".equals(sInteger)) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 轉換整數成為貨幣格式字串
     * @param 整數
     * @return 貨幣格式文字
     */
    public static String toCurrencyString(int iNumber) {
        NumberFormat formatter = new DecimalFormat("#,###,###,###,###,###");
        return formatter.format(iNumber);
    }

    /**
     * 轉換浮點數成為百分比格式字串
     * @param 浮點數
     * @return 百分比格式字串
     */
    public static String toPercentString(float lPercentage) {
        NumberFormat percent = new DecimalFormat("0.0#%");
        return percent.format(lPercentage);
    }

    /**
     * 測試字串是否符合正則表示式(Regular Expression)，若正確回傳true
     * @param 正則表示式
     * @param 測試字串
     * @return true/false
     */
    public static boolean evalRegExpr(String sPattern, String sData){
        Pattern pattern = Pattern.compile(sPattern);
        Matcher matcher = pattern.matcher(sData);
        return matcher.matches();
    }


    /**
     * 從本文中取得兩個標籤中的文字
     * @param 開始標籤
     * @param 結束標籤
     * @param 本文
     * @return 取出之文字
     */
    public static String parseTag(String sStartTag, String sEndTag, String sSourceText) {
        int ns = sStartTag.length();
        // int ne = endTag.length();
        int startIndex = 0;
        int endIndex = 0;
        int valueStart;
        if (sSourceText.indexOf(sStartTag) > 0 && sSourceText.indexOf(sEndTag) > 0) {
            startIndex = sSourceText.indexOf(sStartTag);
            endIndex = sSourceText.indexOf(sEndTag);
            valueStart = startIndex + ns;
        } else {
            return "null";
        }
        String result = sSourceText.substring(valueStart, endIndex);
        return result;
    }

    public static String convertBRtoString(BufferedReader sInput) throws IOException {
        String newLine = System.getProperty("line.separator");
        StringBuilder result = new StringBuilder();
        String line;
        boolean flag = false;
        while ((line = sInput.readLine()) != null) {
            result.append(flag ? newLine : "").append(line);
            flag = true;
        }
        return result.toString();
    }

    /**
     * 轉換文字成為BufferReader
     * @param 文字
     * @return BufferReader
     */
    public static BufferedReader toBufferReader(String sInput) {
        return new BufferedReader(new StringReader(sInput));
    }


    public static String getHttpParameter(HttpServletRequest req, String key) {
        String value = req.getParameter(key);
        if (value == null)
            return "";
        else
            return value;
    }


    /**
     * hex to string
     *
     * @param hex
     * @return
     */
    public static String hexToASCII(String hexValue) {
        StringBuilder output = new StringBuilder("");
        for (int i = 0; i < hexValue.length(); i += 2) {
            String str = hexValue.substring(i, i + 2);
            if (Integer.parseInt(str, 16) != 0)
                output.append((char) Integer.parseInt(str, 16));
        }
        return output.toString();
    }

    /**
     * ASCII To Hex
     *
     * @param asciiValue
     * @return
     */
    public static String asciiToHex(String asciiValue) {
        char[] chars = asciiValue.toCharArray();
        StringBuffer hex = new StringBuffer();
        for (int i = 0; i < chars.length; i++) {
            hex.append(Integer.toHexString((int) chars[i])).append("00");
        }
        return hex.toString().toUpperCase();
    }


    //Trim leading and tailing space
    public static String trim(String source) {
        if (source == null) {
            return "";
        } else {
            return source.trim();
        }

    }

    //trim left zero from alphanumeric text
    public static String trimLeadingZeros(String source) {

        if (TEXT.isEmpty(source))
            return source;

        int length = source.length();
        if (length < 2)
            return source;
        int i;
        for (i = 0; i < length-1; i++) {
            char c = source.charAt(i);
            if (c != '0')
                break;
        }

        if (i == 0)
            return source;

        return source.substring(i);
    }

    /**
     * 自X.500目錄服務取得CN字串
     * @param x500目錄字串
     * @return CN值
     */
    public static String parseCN(String x500) {
        return parseX500(x500, "CN");
    }

    /**
     * 處理X.500目錄字串，取得特定數值
     * @param x500目錄字串
     * @param name值
     * @return value值
     */
    public static String parseX500(String x500, String name) {
        int start = 0;
        int end = 0;
        int len = 0;

        if (x500 == null || name == null)
            return x500;

        len = name.length() + 1;

        if ((start = x500.indexOf(name + "=")) == -1)
            return x500;
        if ((end = x500.indexOf(",", start)) == -1)
            return x500.substring(start + len);
        return x500.substring(start + len, end);
    }



    /**
     * 取得JavaScript跳轉網址的程式語法，用在網頁處理完後轉換至下一頁
     * @param 跳轉去的網址
     * @return Javascript語法字串
     */
    public static String getJSredirect(String sURL) {
        return "<script language=\"javascript\" type=\"text/javascript\">\n" + "<!--\n" + "self.location = \"" + sURL + "\";\n-->\n</script>";
    }

    /**
     * 取得JavaScript跳轉網址的程式語法，用在網頁處理完後轉換至下一頁
     * 跳轉前顯示告警訊息
     * @param 告警訊息字串
     * @param 跳轉去的網址
     * @return Javascript語法字串
     */
    public static String getJSredirect(String sError, String sURL) {
        return "<script language=\"javascript\" type=\"text/javascript\">\n" + "<!--\nalert(\"" + sError + "\");\nself.location = \"" + sURL
                + "\";\n-->\n</script>";
    }

    /**
     * 取得JavaScript跳轉最上層網址的程式語法，用在網頁處理完後轉換至下一頁
     * @param 跳轉去的網址
     * @return Javascript語法字串
     */
    public static String getJSredirectParent(String sURL) {
        return "<script language=\"javascript\" type=\"text/javascript\">\n" + "window.opener.location =\'" + sURL
                + "\';window.parent.focus();window.close();</script>";
    }

    /**
     * 取得JavaScript跳轉最上層網址的程式語法，用在網頁處理完後轉換至下一頁
     * 跳轉前顯示告警訊息
     * @param 告警訊息字串
     * @param 跳轉去的網址
     * @return Javascript語法字串
     */
    public static String getJSredirectParent(String sError, String sURL) {
        return "<script language=\"javascript\" type=\"text/javascript\">\n" + "<!--\nalert(\"" + sError + "\");\nwindow.opener.location =\'" + sURL
                + "\';window.parent.focus();window.close();</script>";
    }

    /**
     * 移除HTML的標籤
     * @param str
     * @param trString
     * @return
     */
    public static String removeHTMLTag(String sInput, String trString) {
        if (sInput == null)
            return "";

        String ret = replaceAll(sInput, trString, "\n");
        ret = replaceAll(ret, "</td>", ",");
        ret = replaceAll(ret, "<td>", "");
        ret = replaceAll(ret, "</tr>", "");
        ret = replaceAll(ret, "&nsbp", "");
        return ret;
    }

    /**
     * 取代字串內所有文字
     * @param 原始字串
     * @param 舊文字
     * @param 新文字
     * @return 取代後字串
     */
    public static String replaceAll(String sOrigin, String sOld, String sNew) {
        StringBuffer stringBuffer = new StringBuffer(sOrigin);
        int index = sOrigin.length();
        int offset = sOld.length();

        while ((index = sOrigin.lastIndexOf(sOld, index - 1)) > -1) {
            stringBuffer.replace(index, index + offset, sNew);
        }
        return stringBuffer.toString();
    }


}
