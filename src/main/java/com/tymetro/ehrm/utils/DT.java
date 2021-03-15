/*****************************************************************************************
 Class Name: DT
 Package: com.kjsoft.common
 Purpose: DataTable Utility
 Memo:
 History:
 Mod #	Date		By			Description
 -------	----------	--------	--------------------------------------------------
 *000    2019-05-07 	Jimmy Lee	Initial Code

 Copyright (c) 2019, KJSOFT. All rights reserved.
 *****************************************************************************************/

package com.tymetro.ehrm.utils;

import java.sql.SQLException;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;


import com.google.gson.JsonObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 處理DataTable在Server Side Processing
 * @author Jimmy Lee ( <a href="mailto:jimmy@kjsoft.com.tw">jimmy@kjsoft.com.tw</a> )
 * @version 1.0
 */
public class DT {

    private static final Logger logger = LoggerFactory.getLogger(DT.class);

    /**
     * 取得DataTables資料回傳之jSON字串
     * @throws SQLException
     */
    public static String getDataTableRespond(HttpServletRequest hRequest, String sInnerSQL) throws SQLException {
        if (DB.isOracle()) {
            return getDataTableRespondOracle(hRequest, sInnerSQL);
        }
        return "";
    }

    /**
     * 取得DataTables資料回傳之jSON字串
     * @throws SQLException
     */
    private static String getDataTableRespondOracle(HttpServletRequest hRequest, String sInnerSQL) throws SQLException {

        String return_datatable_json = "";

        try {

            // 取得並分辨Client查詢參數
            DTParam param = new DTParam(hRequest);

            // 組搜尋SQL語法
            String searchSql = "";
            if(!TEXT.isEmpty(param.sSearchValue) && param.iColumnCount > 0 ) {
                String[] columns = param.sColumnData;
                for( int i = 0; i < columns.length; i++ ) {
                    if( param.bColumnSearchable[i] ) {
                        searchSql += (searchSql.length() == 0 ? " where (" : " or ") + columns[i] + " like '%" + param.sSearchValue + "%'";
                    }
                }
                if( searchSql.length() > 0 ) {
                    searchSql += ")";
                }
            }

            // 下SQL計算資料筆數
            int iRecordsTotal; // total number of records (unfiltered)
            int iRecordsFiltered; //value will be set when code filters companies by keyword
            iRecordsTotal = TEXT.toInt(DB.queryScalar("select count(*) as COUNT from (" + sInnerSQL + ")"));
            if( TEXT.isEmpty(searchSql) ) {
                iRecordsFiltered = TEXT.toInt(DB.queryScalar("select count(*) as COUNT from (" + sInnerSQL + ") " + searchSql));
            } else {
                iRecordsFiltered = iRecordsTotal;
            }

            // 組排序SQL語法
            String orderSql = "";
            if (!TEXT.isNull(param.sOrderColumn)) {
                orderSql += " order by " + param.sOrderColumn + " " + param.sOrderDir + " NULLS LAST";
            }
//舊的邏輯，可進行多欄位排序
//			orderSql += (param.iSortingCols > 0 ? " order by " : "");
//			for( int i = 0; i < param.iSortingCols; i++ ) {
//				int sortBy = param.iSortCol[i];
//				String sortField = columns[sortBy];
//				if( param.bSortable[sortBy] ) {
//					orderSql += sortField + " " + param.sSortDir[i] + " NULLS LAST" + (i<param.iSortingCols-1?",":"");
//				}
//			}

            // 組完整SQL抓資料
            String sql = "select * from (select t.*, ROWNUM as rnum from (" + sInnerSQL + orderSql + ") t"  + searchSql + ") where rnum between "+(param.iStart+1)+" and "+(param.iStart+param.iLength);
            List<Map<String, Object>> list = DB.query(sql);
//			for( Map<String, Object> map : list ) {
//				if( map.containsKey("dt_rowid") ) {
//					map.put("DT_RowId", map.get("dt_rowid"));
//					map.remove("dt_rowid");
//				}
//			}

            // 建立json物件
            JsonObject jsonResponse = new JsonObject();
            jsonResponse.addProperty("draw", param.iDraw);
            jsonResponse.addProperty("recordsTotal", iRecordsTotal);
            jsonResponse.addProperty("recordsFiltered", iRecordsFiltered);
            jsonResponse.add("data", JSON.toJson(list));

            // 回傳json字串
            return_datatable_json = jsonResponse.toString();

        } catch (Exception e) {
            e.printStackTrace();
        }
        return return_datatable_json;

    }

}
