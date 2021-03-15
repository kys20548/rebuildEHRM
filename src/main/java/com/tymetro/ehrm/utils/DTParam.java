/*****************************************************************************************
 Class Name: DTParam
 Package: com.kjsoft.common
 Purpose: DataTable Parameter Object
 Memo:
 History:
 Mod #	Date		By			Description
 -------	----------	--------	--------------------------------------------------
 *000    2019-05-07 	Jimmy Lee	Initial Code

 Copyright (c) 2019, KJSOFT. All rights reserved.
 *****************************************************************************************/

package com.tymetro.ehrm.utils;

import javax.servlet.http.HttpServletRequest;

/**
 * 處理DataTable傳入參數
 * @author Jimmy Lee ( <a href="mailto:jimmy@kjsoft.com.tw">jimmy@kjsoft.com.tw</a> )
 * @version 1.0
 */
public class DTParam {

    // Request sequence number sent by DataTable, same value must be returned in

    // This is used by DataTables to ensure that the Ajax returns from server-side processing requests are drawn in sequence by DataTables.
    public int iDraw = 0;

    // Paging first record indicator. This is the start point in the current data set (0 index based - i.e. 0 is the first record).
    public int iStart = 0;

    // Number of records that the table can display in the current draw. Note that this can be -1 to indicate that all records should be returned (although that negates any benefits of server-side processing!).
    public int iLength = 0;

    // Global search value. To be applied to all columns which have searchable as true.
    public String sSearchValue;

    // true if the global filter should be treated as a regular expression for advanced searching, false otherwise.
    public boolean bSearchRegex;

    // Column to which ordering should be applied. This is an index reference to the columns array of information that is also submitted to the server.
    public int iOrderColumn;
    public String sOrderColumn;

    // Column to which ordering should be applied. This is an index reference to the columns array of information that is also submitted to the server.
    public String sOrderDir;

    // Column Number.
    public int iColumnCount;

    // Column's data source.
    public String[] sColumnData;

    // Column's name.
    public String[] sColumnName;

    // Flag to indicate if this column is searchable.
    public boolean[] bColumnSearchable;

    // Flag to indicate if this column is orderable.
    public boolean[] bColumnOrderable;

    // Flag to indicate if this column is orderable.
    public String[] sColumnSearchValue;

    // Flag to indicate if this column is orderable.
    public boolean[] bColumnSearchRegex;


    //old framework prior
//	public String sEcho;
//	public String sSearchKeyword;
//	public boolean bRegexKeyword;
//	public int iDisplayLength;
//	public int iDisplayStart;
//	public int iColumns;
//	public String[] sSearch;
//	public boolean[] bSearchable;
//	public boolean[] bSortable;
//	public boolean[] bRegex;
//	public int iSortingCols;
//	public String[] sSortDir;
//	public int[] iSortCol;
//	public String sColumns;

    public DTParam(HttpServletRequest hRequest) {
        this.parse(hRequest);
    }

    //分析datatable回傳參數
    public void parse(HttpServletRequest hRequest) {

        if (!TEXT.isEmpty(hRequest.getParameter("draw"))) {
            iDraw = TEXT.toInt(hRequest.getParameter("draw"));
            iStart =  TEXT.toInt(hRequest.getParameter("start"));
            iLength =  TEXT.toInt(hRequest.getParameter("length"));
            sSearchValue = hRequest.getParameter("search[value]");
            sSearchValue = hRequest.getParameter("search[value]");
            bSearchRegex = "true".equals(hRequest.getParameter("length")) ? true : false ;
            sOrderDir = hRequest.getParameter("order[0][dir]");
            if (!TEXT.isEmpty(hRequest.getParameter("order[0][column]"))) {
                iOrderColumn = TEXT.toInt(hRequest.getParameter("order[0][column]"));
                sOrderColumn = hRequest.getParameter("columns["+iOrderColumn+"][data]");
            } else {
                iOrderColumn = 0;
                sOrderColumn = "";
            }
            //計算出欄位數量
            for (iColumnCount = 0; ; iColumnCount++) {
                if (TEXT.isNull(hRequest.getParameter("columns["+iColumnCount+"][data]"))) {
                    break;
                }
            }
            //初始化資料欄位參數
            sColumnData = new String[iColumnCount];
            sColumnName = new String[iColumnCount];
            bColumnSearchable = new boolean[iColumnCount];
            bColumnOrderable = new boolean[iColumnCount];
            sColumnSearchValue = new String[iColumnCount];
            bColumnSearchRegex = new boolean[iColumnCount];
            //填入資料欄位參數
            for (int i = 0; i < iColumnCount; i++) {
                sColumnData[i] = hRequest.getParameter("columns["+i+"][data]");
                sColumnName[i] = hRequest.getParameter("columns["+i+"][name]");
                bColumnSearchable[i] = TEXT.toBoolean(hRequest.getParameter("columns["+i+"][searchable]"));
                bColumnOrderable[i] = TEXT.toBoolean(hRequest.getParameter("columns["+i+"][orderable]"));
                sColumnSearchValue[i] = hRequest.getParameter("columns["+i+"][search][value]");
                bColumnSearchRegex[i] = TEXT.toBoolean(hRequest.getParameter("columns["+i+"][search][regex]"));
                if (TEXT.toInt(sColumnData[i]) == i) {
                    bColumnSearchable[i] = false;
                }
            }

        } else {
            iDraw = 0;
        }

        //old framework
//		if ((request.getParameter("sEcho") != null) && (request.getParameter("sEcho") != "")) {
//			sEcho = request.getParameter("sEcho");
//			sSearchKeyword = request.getParameter("sSearch");
//			bRegexKeyword = Boolean.parseBoolean(request.getParameter("bRegex"));
//			sColumns = request.getParameter("sColumns");
//			iDisplayStart = Integer.parseInt(request.getParameter("iDisplayStart"));
//			iDisplayLength = Integer.parseInt(request.getParameter("iDisplayLength"));
//			iColumns = Integer.parseInt(request.getParameter("iColumns"));
//			sSearch = new String[iColumns];
//			bSearchable = new boolean[iColumns];
//			bSortable = new boolean[iColumns];
//			bRegex = new boolean[iColumns];
//			for (int i = 0; i < iColumns; i++) {
//				sSearch[i] = request.getParameter("sSearch_" + i);
//				bSearchable[i] = Boolean.parseBoolean(request.getParameter("bSearchable_" + i));
//				bSortable[i] = Boolean.parseBoolean(request.getParameter("bSortable_" + i));
//				bRegex[i] = Boolean.parseBoolean(request.getParameter("bRegex_" + i));
//			}
//
//			iSortingCols = Integer.parseInt(request.getParameter("iSortingCols"));
//			sSortDir = new String[iSortingCols];
//			iSortCol = new int[iSortingCols];
//			for (int i = 0; i < iSortingCols; i++) {
//				sSortDir[i] = request.getParameter("sSortDir_" + i);
//				iSortCol[i] = Integer.parseInt(request.getParameter("iSortCol_" + i));
//			}
//		}

    }
}
