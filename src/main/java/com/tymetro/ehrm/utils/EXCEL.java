package com.tymetro.ehrm.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import static com.tymetro.ehrm.utils.CHECKUTIL.nonNullBigDecimal;

public class EXCEL {

    public static XSSFWorkbook createExcel(XSSFWorkbook workbook, List<Map<String, Object>> list, String path) {

        CellStyle style1 = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        fontStyle1.setFontName("標楷體");
        fontStyle1.setFontHeightInPoints((short) 16);
        fontStyle1.setBold(true);
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setFont(fontStyle1);

        CellStyle style2 = workbook.createCellStyle();
        Font fontStyle2 = workbook.createFont();
        fontStyle2.setFontName("標楷體");
        fontStyle2.setFontHeightInPoints((short) 13);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setFont(fontStyle2);

        CellStyle style3 = workbook.createCellStyle();
        Font fontStyle3 = workbook.createFont();
        fontStyle3.setFontName("標楷體");
        fontStyle3.setFontHeightInPoints((short) 12);
        style3.setAlignment(HorizontalAlignment.CENTER);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);
        style3.setBorderBottom(BorderStyle.THIN);
        style3.setBorderLeft(BorderStyle.THIN);
        style3.setBorderRight(BorderStyle.THIN);
        style3.setBorderTop(BorderStyle.THIN);
        style3.setFont(fontStyle3);

        CellStyle style4 = workbook.createCellStyle();
        Font fontStyle4 = workbook.createFont();
        fontStyle4.setFontName("標楷體");
        fontStyle4.setFontHeightInPoints((short) 12);
        fontStyle4.setColor(IndexedColors.RED.getIndex());
        style4.setAlignment(HorizontalAlignment.CENTER);
        style4.setVerticalAlignment(VerticalAlignment.CENTER);
        style4.setBorderBottom(BorderStyle.THIN);
        style4.setBorderLeft(BorderStyle.THIN);
        style4.setBorderRight(BorderStyle.THIN);
        style4.setBorderTop(BorderStyle.THIN);
        style4.setFont(fontStyle4);

        int rows = 0;
        XSSFSheet sheet = null;
        XSSFRow xssfRow;
        XSSFCell cell;

        sheet = workbook.createSheet("保額差異");
        // 8欄位
        int[] cellWidth = { 2750, 2750,2750, 2750, 2750, 2750, 2750, 2750, 2750};
        for (int i = 0; i < cellWidth.length; i++) {
            sheet.setColumnWidth(i, cellWidth[i]);
        }
        //header
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 600);
        cell = xssfRow.createCell(0);
        String titleYear =Integer.parseInt(((String)list.get(0).get("INSURANCEYYMM")).substring(0, 4))-1911+"年";
        String titleMonth =((String)list.get(0).get("INSURANCEYYMM")).substring(4,6)+"月";
        cell.setCellValue("桃園大眾捷運股份有限公司"+titleYear+titleMonth+"保額差異比對");
        cell.setCellStyle(style1);
        for (int i = 1; i < cellWidth.length; i++) {
            cell = xssfRow.createCell(i);
            cell.setCellStyle(style1);
        }
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 0, cellWidth.length - 1));
        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        cell = xssfRow.createCell(0);
        cell.setCellValue("投保年月");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(1);
        cell.setCellValue("員工編號");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(2);
        cell.setCellValue("員工姓名");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(3);
        cell.setCellValue("本月勞保");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(4);
        cell.setCellValue("前月勞保");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(5);
        cell.setCellValue("本月健保");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(6);
        cell.setCellValue("前月健保");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(7);
        cell.setCellValue("本月勞退");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(8);
        cell.setCellValue("前月勞退");
        cell.setCellStyle(style2);
        String year="";
        String month="";
        for (Map map : list) {
            rows = rows + 1;
            xssfRow = sheet.createRow(rows);

            cell = xssfRow.createCell(0);
            year =Integer.parseInt(((String)list.get(0).get("INSURANCEYYMM")).substring(0, 4))-1911+"/";
            month =((String)list.get(0).get("INSURANCEYYMM")).substring(4);
            cell.setCellValue(year+month);
            cell.setCellStyle(style3);

            cell = xssfRow.createCell(1);
            cell.setCellValue((String) map.get("EMPNO"));
            cell.setCellStyle(style3);

            cell = xssfRow.createCell(2);
            cell.setCellValue((String) map.get("EMPNAME"));
            cell.setCellStyle(style3);

            cell = xssfRow.createCell(3);
            cell.setCellValue((nonNullBigDecimal(map.get("LABOR_INSURANCE"))).doubleValue());
            if((nonNullBigDecimal(map.get("LABOR_INSURANCE")).equals(nonNullBigDecimal(map.get("LAST_LABOR_INSURANCE"))))) {
                cell.setCellStyle(style3);
            }else {
                cell.setCellStyle(style4);
            }

            cell = xssfRow.createCell(4);
            cell.setCellValue((nonNullBigDecimal(map.get("LAST_LABOR_INSURANCE"))).doubleValue());
            cell.setCellStyle(style3);

            cell = xssfRow.createCell(5);
            cell.setCellValue((nonNullBigDecimal( map.get("HEALTH_INSURANCE"))).doubleValue());
            if((nonNullBigDecimal(map.get("HEALTH_INSURANCE")).equals(nonNullBigDecimal(map.get("LAST_HEALTH_INSURANCE"))))) {
                cell.setCellStyle(style3);
            }else {
                cell.setCellStyle(style4);
            }

            cell = xssfRow.createCell(6);
            cell.setCellValue((nonNullBigDecimal( map.get("LAST_HEALTH_INSURANCE"))).doubleValue());
            cell.setCellStyle(style3);

            cell = xssfRow.createCell(7);
            cell.setCellValue((nonNullBigDecimal( map.get("RETIRE_INSURANCE"))).doubleValue());
            if((nonNullBigDecimal(map.get("RETIRE_INSURANCE")).equals(nonNullBigDecimal(map.get("LAST_RETIRE_INSURANCE"))))) {
                cell.setCellStyle(style3);
            }else {
                cell.setCellStyle(style4);
            }

            cell = xssfRow.createCell(8);
            cell.setCellValue((nonNullBigDecimal( map.get("LAST_RETIRE_INSURANCE"))).doubleValue());
            cell.setCellStyle(style3);
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return workbook;
    }
}
