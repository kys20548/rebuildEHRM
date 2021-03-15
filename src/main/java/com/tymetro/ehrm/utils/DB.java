/*****************************************************************************************
 Class Name: DB
 Package: com.kjsoft.common
 Purpose: Database Utility
 Memo:
 History:
 Mod #	Date		By			Description
 -------	----------	--------	--------------------------------------------------
 *000    2019-05-07 	Jimmy Lee	Initial Code

 Copyright (c) 2019, KJSOFT. All rights reserved.
 *****************************************************************************************/

package com.tymetro.ehrm.utils;

import java.sql.Connection;
import java.sql.SQLException;
import java.util.List;
import java.util.Map;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.support.rowset.SqlRowSet;
import org.springframework.jdbc.support.rowset.SqlRowSetMetaData;

/**
 * 資料庫函數工具，存取資料庫資料
 * @author Yung Lee ( <a href="mailto:yung@kjsoft.com.tw">yung@kjsoft.com.tw</a> )
 * @version 1.0
 */
public class DB {

    //private static final Logger logger = Logger.getLogger(DB.class);

    // JDBC資料來源物件
    private static JdbcTemplate jdbcTemplate;
    private static String dbname;

    /**
     * 載入JdbcTemplate物件，供INIT元件叫用
     * @param jJdbcTemplate物件
     */
    public static void setJdbcTemplate(JdbcTemplate jJdbcTemplate) {
        jdbcTemplate = jJdbcTemplate;
        jdbcTemplate.setFetchSize(500);
    }

    /**
     * 取得JdbcTemplate物件
     * @return JdbcTemplate物件
     */
    public static JdbcTemplate getJdbcTemplate() {
        return jdbcTemplate;
    }

    /**
     * SQL語法更新資料表欄位資訊
     * @param SQL語法
     * @return 成功筆數
     * @throws SQLException
     */
    public static int update(String sSQL) throws SQLException {
        int ret = 0;
        //logger.debug("DB.update: "+sSQL);
        ret = jdbcTemplate.update(sSQL);
        //logger.debug("DB.update: record="+ret);
        return ret;
    }

    /**
     *  SQL語法讀取資料
     * @param SQL語法
     * @return 集合物件
     * @throws SQLException
     */
    public static List<Map<String, Object>> query(String sSQL) throws SQLException {
        //logger.debug("DB.query: "+sSQL);
        List<Map<String, Object>> result = jdbcTemplate.queryForList(sSQL);
        return result;
    }

    /**
     * SQL語法(含參數)讀取資料
     * @param sSQL SQL語法
     * @param oArg 查詢參數
     * @return 集合物件
     * @throws SQLException
     */
    public static List<Map<String, Object>> query(String sSQL, Object[] oArg) throws SQLException {
        //logger.debug("DB.query: "+sSQL);
        List<Map<String, Object>> result = jdbcTemplate.queryForList(sSQL, oArg);
        return result;
    }

    /**
     *  SQL語法讀取RowSet
     * @param SQL語法
     * @return RowSet
     * @throws SQLException
     */
    public static SqlRowSet queryForRowSet(String sSQL) throws SQLException {
        //logger.debug("DB.queryForRowSet: "+sSQL);
        return jdbcTemplate.queryForRowSet(sSQL);
    }

    /**
     * SQL語法(含參數)讀取RowSet
     * @param sSQL SQL語法
     * @param oArg 查詢參數
     * @return RowSet
     * @throws SQLException
     */
    public static SqlRowSet queryForRowSet(String sSQL, Object[] oArg) throws SQLException {
        //logger.debug("DB.queryForRowSet: "+sSQL);
        return jdbcTemplate.queryForRowSet(sSQL, oArg);
    }

    /**
     * 讀取SQL回傳值第一行第一列之字串值
     * @param SQL語法
     * @return 第一行第一列之字串值
     * @throws SQLException
     */
    public static String queryScalar(String sSQL) throws SQLException {
        //logger.debug("DB.getQueryScalar: "+sSQL);
        return (String) jdbcTemplate.queryForObject(sSQL, String.class);
    }

    /**
     * 讀取SQL回傳值第一行第一列之字串值
     * @param SQL語法
     * @param 查詢參數
     * @return 第一行第一列之字串值
     * @throws SQLException
     */
    public static String queryScalar(String sSQL, Object[] oArg) throws SQLException {
        //logger.debug("DB.getQueryScalar: "+sSQL+", argument: "+oArg.toString());
        return (String) jdbcTemplate.queryForObject(sSQL, oArg, String.class);
    }

    /**
     * 將SQL查詢字串轉換成HTML表格
     * @param sSQL 資料庫查詢SQL
     * @return HTML表格
     * @throws Exception
     */
    public static String getHTMLTable(String sSQL) throws Exception {
        return getHTMLTable(sSQL, "", "", "", "", true);
    }

    /**
     * 將SQL查詢字串轉換成HTML表格
     * @param sSQL 資料庫查詢SQL
     * @param sTB TABLE表格之TB內HTML屬性
     * @param sTH TABLE表格之TH內HTML屬性
     * @param sTR TABLE表格之TR內HTML屬性
     * @param sTD TABLE表格之TD內HTML屬性
     * @param bShowCount 是否顯示表格尾之加總數據
     * @return HTML表格
     * @throws Exception
     */
    public static String getHTMLTable(String sSQL, String sTB, String sTH, String sTR, String sTD, Boolean bShowCount) throws Exception {
        SqlRowSet srs = null;
        SqlRowSetMetaData srsmd = null;
        int numberOfColumns = 0;
        int rowCounter = 0;
        StringBuffer output = new StringBuffer();

        try {
            //logger.debug("DB.getHTMLTable SQL=" + sSQL);

            srs = jdbcTemplate.queryForRowSet(sSQL);

            // 建立表格<table>
            if (TEXT.isEmpty(sTB))
                output.append("<table border=1 cellspacing=0 cellpadding=0>\n");
            else
                output.append("<table ").append(sTB).append(">\n");

            // 建立表頭<th>
            srsmd = srs.getMetaData();
            numberOfColumns = srsmd.getColumnCount();
            output.append("<tr>\n");
            output.append("<th ").append(sTH).append(">").append("#").append("</th>\n");
            for (int i = 1; i <= numberOfColumns; i++) {
                output.append("<th ").append(sTH).append(">").append(srsmd.getColumnName(i)).append("</th>\n");
            }
            output.append("</tr>\n");

            // 建立表列<tr>
            while (srs.next()) {
                rowCounter++;
                output.append("<tr ").append(sTR).append(">\n");
                output.append("<td ").append(sTD).append(">").append(rowCounter).append("</td>\n");
                for (int i = 1; i <= numberOfColumns; i++) {
                    String field = srs.getString(i);
                    field = (TEXT.isEmpty(field) ? "&nbsp;" : field);
                    output.append("<td ").append(sTD).append(">").append(field).append("</td>\n");
                }
                output.append("</tr>\n");
            }

            // 顯示表尾資料筆數
            if (bShowCount) {
                // 無資料時回傳
                if (rowCounter == 0) {
                    output.append("<tr ").append(sTR).append("><td ")
                            .append(sTD).append(" colspan=\"")
                            .append(numberOfColumns+1)
                            .append("\"><strong>查無資料!!!</strong></td></tr>\n");
                } else {
                    output.append("<tr ").append(sTR).append("><td ")
                            .append(sTD).append(" colspan=\"")
                            .append(numberOfColumns+1).append("\">資料筆數共計：")
                            .append(rowCounter).append(" 筆</td></tr>\n");
                }
            }
            output.append("</table>\n");
        } catch (Exception e) {
            //logger.error("DB.getHTMLTable FAILED,REASON="+e.toString());
            //throw EXCP.getException(1000, "資料庫讀取錯誤", e);
        }
        return output.toString();
    }




    /**
     * 檢查資料庫狀態
     * @return true/false
     */
    public static boolean isAlive() {
        boolean check_result = false;
        try {
            jdbcTemplate.queryForList("SELECT 1");
            check_result = true;
        } catch (Exception e) {
            check_result = false;
        }
        return check_result;
    }

    /**
     * 檢查資料庫名稱
     * @return DB NAME
     */
    public static String getDBName() {
        Connection conn = null;
        if (!TEXT.isEmpty(dbname)) {
            return dbname;
        }
        try {
            conn = jdbcTemplate.getDataSource().getConnection();
            dbname = conn.getMetaData().getDatabaseProductName();
        } catch (Exception e) {
            return "";
        } finally {
            try {
                if (conn != null && !conn.isClosed()) conn.close();
            } catch (Exception e) {}
        }
        return dbname;
    }

    /**
     * 檢查資料庫是否為Oracle
     * @return true/false
     */
    public static boolean isOracle() {
        return "Oracle".equals(getDBName());
    }

    /**
     * 檢查資料庫是否為MSSQL
     * @return true/false
     */
    public static boolean isMSSQL() {
        return "Microsoft SQL Server".equals(getDBName());
    }

    /**
     * 檢查資料庫是否為SQLite
     * @return true/false
     */
    public static boolean isSQLite() {
        return "SQLite".equals(getDBName());
    }

}
