/*****************************************************************************************
 Class Name: JSON
 Package: com.kjsoft.common
 Purpose: Json Data Utility
 Memo:
 History:
 Mod #	Date		By			Description
 -------	----------	--------	--------------------------------------------------
 *000    2019-05-07 	Jimmy Lee	Initial Code

 Copyright (c) 2019, KJSOFT. All rights reserved.
 *****************************************************************************************/

package com.tymetro.ehrm.utils;

import java.lang.reflect.Type;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import com.google.gson.reflect.TypeToken;

/**
 * 處理Json資料格式
 * @author Jimmy Lee ( <a href="mailto:jimmy@kjsoft.com.tw">jimmy@kjsoft.com.tw</a> )
 * @version 1.0
 */
public class JSON {

    private static Gson gson = new GsonBuilder().setDateFormat("yyyy-MM-dd HH:mm:ss").serializeNulls().create();

    /**
     * 轉換ListMap資料結構成為JsonArray
     * @param LispMap
     * @return JsonArray
     */
    public static JsonArray toJson(List<Map<String, Object>> lmData) {
        return gson.toJsonTree(lmData).getAsJsonArray();
    }

    /**
     * 轉換JsonArray資料結構成為ListMap
     * @param JsonArray
     * @return LispMap
     */
    public static List<Map<String, Object>> toListMap(JsonArray jData) {
        Type listType = new TypeToken<List<HashMap<String, Object>>>(){}.getType();
        List<Map<String, Object>> result = gson.fromJson(jData.toString(), listType);
        return result;
    }


    /**
     * 將Json字串格式化成為美觀可輸出顯示之格式
     * @param 原始Json字串
     * @return 格式化後美觀之Json字串
     */
    public static String prettyJson(String sJsonString) {
        Gson gsonp = new GsonBuilder().setPrettyPrinting().create();
        JsonParser jp = new JsonParser();
        JsonElement je = jp.parse(sJsonString);
        String prettyJsonString = gsonp.toJson(je);
        return prettyJsonString;
    }

}
