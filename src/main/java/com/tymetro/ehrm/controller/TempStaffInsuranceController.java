package com.tymetro.ehrm.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.tymetro.ehrm.model.TempStaffInsurance;
import com.tymetro.ehrm.utils.DB;
import com.tymetro.ehrm.utils.DT;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.BatchPreparedStatementSetter;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

@Controller
@RequestMapping(path = "/temp/TempStaffInsuranceController")
public class TempStaffInsuranceController {

//    private final String UPLOADED_FOLDER = CFG.getWebAppRootPath()+File.separator+"upload";
    private final String UPLOADED_FOLDER = "/upload";
    private final String FILE_NAME = "tempInsuExcel.xlsx";

    private int[] healthRange= {24000,25200,26400,27600,28800,30300,31800,33300,34800,36300,38200,40100,42000,43900,45800};
    private int[] laborRange= {11100,12540,13500,15840,16500,17280,17880,19047,20008,21009,22000,23100,23800,24000,25200,26400,27600,28800,30300,31800,33300,34800,36300,38200,40100,42000,43900,45800};
    private int[] retireRange= {1500,3000,4500,6000,7500,8700,9900,11100,12540,13500,15840,16500,17280,17880,19047,20008,21009,22000,23100,23800,24000,25200,26400,27600,28800,30300,31800,33300,34800,36300,38200,40100,42000,43900,45800};

    @Autowired
    ServletContext context;

    @ResponseBody
    @RequestMapping(path = "/tempStaff_insu_list")
    public String source(HttpServletRequest request) throws SQLException {
        String dateYM = request.getParameter("salaryYM");
        if (dateYM == null || dateYM.trim().equals("")) {
            dateYM = DB.queryScalar("select max(INSURANCEYYMM) from TEMP_TEMPSTAFF_INSURANCE");
        }

        String sql = " select INSURANCEYYMM,EMPNO, HOUR_RATE, HOURS, SALARY, LABOR_INSURANCE, HEALTH_INSURANCE, RETIRE_INSURANCE ";
        sql += " from TEMP_TEMPSTAFF_INSURANCE ";
        sql += " where INSURANCEYYMM='" + dateYM + "' ";

        return DT.getDataTableRespond(request, sql);
    }

    @RequestMapping(value = "/upload")
    public String upload(HttpServletRequest request,ModelMap model, @RequestParam("file") MultipartFile file) throws SQLException {
        String userName=(String)request.getSession().getAttribute("tymetro_ehrm_user_name");
        try {
            byte[] bytes = file.getBytes();
            File folder =new File(UPLOADED_FOLDER);
            if(!folder.exists()) {
                folder.mkdir();
            }
            File excelFile =new File(UPLOADED_FOLDER+File.separator+FILE_NAME);
            FileUtils.writeByteArrayToFile(excelFile, bytes);

            List<TempStaffInsurance> datalist = getFileList(UPLOADED_FOLDER+File.separator + FILE_NAME,userName);
            updateBatch(datalist);
            insertBatch(datalist);
            excelFile.delete();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            request.getSession().setAttribute("errorMsg", e.toString());
            return "redirect:/temp/temp_error.jsp";
        }
        return "redirect:/temp/temp_insu.jsp";
    }

    @RequestMapping(value = "/download")
    public void download(HttpServletRequest request, HttpServletResponse response) throws SQLException {
        String salaryYM = request.getParameter("dateYM");
        XSSFWorkbook workbook;
        String sql ="select a.*,b.LABOR_INSURANCE LAST_LABOR_INSURANCE,b.HEALTH_INSURANCE LAST_HEALTH_INSURANCE,b.RETIRE_INSURANCE LAST_RETIRE_INSURANCE ";
        sql += " from (select * from TEMP_TEMPSTAFF_INSURANCE where INSURANCEYYMM='"+salaryYM+"') a,";
        sql += " (select * from TEMP_TEMPSTAFF_INSURANCE where INSURANCEYYMM=to_char( add_months(to_date('"+salaryYM+"','yyyyMM'),-1),'yyyyMM')) b ";
        sql += " where a.EMPNO=b.EMPNO(+) ";
        sql += " order by a.EMPNO asc ";
        List<Map<String, Object>> list = DB.query(sql);
        try {
            File folder =new File(UPLOADED_FOLDER);
            if(!folder.exists()) {
                folder.mkdir();
            }
            String titleYear =(Integer.parseInt(salaryYM.substring(0,4))-1911)+"???";
            String titleMounth = salaryYM.substring(4,6)+"???";
            String fileName="?????????????????????????????????"+titleYear+titleMounth+".xlsx";
            fileName = URLEncoder.encode(fileName,"UTF-8");

            File file = new File(UPLOADED_FOLDER+File.separator + fileName);
            workbook = createExcel(new XSSFWorkbook(), list, UPLOADED_FOLDER+File.separator + fileName);

            if (file.exists()) {
                String mimeType = context.getMimeType(file.getPath());

                if (mimeType == null) {
                    mimeType = "application/octet-stream";
                }

                response.setContentType(mimeType);
                response.addHeader("Content-Disposition", "attachment; filename=" + fileName);
                response.setContentLength((int) file.length());

                OutputStream os = response.getOutputStream();
                FileInputStream fis = new FileInputStream(file);
                byte[] buffer = new byte[4096];
                int b = -1;

                while ((b = fis.read(buffer)) != -1) {
                    os.write(buffer, 0, b);
                }

                fis.close();
                os.close();

            } else {
                System.out.println("Requested " + fileName + " file not found!!");
            }
            workbook.close();
        } catch (IOException e) {
            System.out.println("Error:- " + e.getMessage());
        }
    }

    public XSSFWorkbook createExcel(XSSFWorkbook workbook, List<Map<String, Object>> list, String path) {

        CellStyle style1 = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        fontStyle1.setFontName("?????????");
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
        fontStyle2.setFontName("?????????");
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
        fontStyle3.setFontName("?????????");
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
        fontStyle4.setFontName("?????????");
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

        sheet = workbook.createSheet("????????????");
        // 8??????
        int[] cellWidth = { 2750, 2750,2750, 2750, 2750, 2750, 2750, 2750, 2750};
        for (int i = 0; i < cellWidth.length; i++) {
            sheet.setColumnWidth(i, cellWidth[i]);
        }
        //header
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 600);
        cell = xssfRow.createCell(0);
        String titleYear =Integer.parseInt(((String)list.get(0).get("INSURANCEYYMM")).substring(0, 4))-1911+"???";
        String titleMonth =((String)list.get(0).get("INSURANCEYYMM")).substring(4,6)+"???";
        cell.setCellValue("????????????????????????????????????"+titleYear+titleMonth+"??????????????????");
        cell.setCellStyle(style1);
        for (int i = 1; i < cellWidth.length; i++) {
            cell = xssfRow.createCell(i);
            cell.setCellStyle(style1);
        }
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 0, cellWidth.length - 1));
        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        cell = xssfRow.createCell(0);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(1);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(2);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(3);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(4);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(5);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(6);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(7);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(8);
        cell.setCellValue("????????????");
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

    @SuppressWarnings({ "rawtypes", "unchecked" })
    public List<TempStaffInsurance> getFileList(String filePath, String userName) throws Exception {

        List<List<Object>> excelList = new ArrayList();

        Workbook workbook;
        try {
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            iterator.next();
            while (iterator.hasNext()) {
                List list = new ArrayList();

                Row currentRow = iterator.next();
                //for (int i = 0; i < currentRow.getLastCellNum(); i++) {
                for (int i = 0; i <= 5; i++) {
                    if (currentRow.getCell(i) == null) {
                        list.add("");
                    } else if (currentRow.getCell(i).getCellTypeEnum() == CellType.STRING) {
                        list.add(currentRow.getCell(i).getStringCellValue());
                    } else if (currentRow.getCell(i).getCellTypeEnum() == CellType.NUMERIC) {
                        list.add(currentRow.getCell(i).getNumericCellValue());
                    } else if(currentRow.getCell(i).getCellTypeEnum()==CellType.BLANK){
                        list.add("");
                    }else {
                        System.out.println("???????????????");
                    }

                }
                excelList.add(list);
            }

            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        List<TempStaffInsurance> rtList = new ArrayList();
        for (List list : excelList) {
            TempStaffInsurance tsi = new TempStaffInsurance();
            // ????????????
            try {
                tsi.setInsuranceYyMm((Integer.valueOf(((String) list.get(0)).substring(0, 3)) + 1911)+ ((String) list.get(0)).substring(4));
                tsi.setEmpNo((String) list.get(1));// ????????????
                tsi.setEmpName((String) list.get(2));// ????????????
                tsi.setHourRate("".equals(list.get(3))==true?0:(Double)list.get(3));
                tsi.setHours("".equals(list.get(4))==true?0:(Double)list.get(4));
                tsi.setSalary("".equals(list.get(5))==true?0:((Double)list.get(5)).intValue());
            } catch (Exception e) {
                throw new Exception((String)list.get(2)+"  ???????????????????????????????????????");
            }

            tsi.setUpdateDate(new Date());
            tsi.setUpdateUser(userName);

            /*
             * ??????????????????
             *
             */

            //???????????? CAL_SALARY
            BigDecimal bd = BigDecimal.ZERO;

            if(tsi.getSalary()==0) {
                bd = new BigDecimal(tsi.getHourRate()*tsi.getHours());
                bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
                tsi.setSalary(bd.intValue());
            }

            for(int laborinsu:laborRange) {
                if(tsi.getSalary()<=laborinsu) {
                    tsi.setLaborInsurance(laborinsu);
                    break;
                }
            }

            for(int laborinsu:healthRange) {
                if(tsi.getSalary()<=laborinsu) {
                    tsi.setHealthInsurance(laborinsu);
                    break;
                }
            }

            for(int laborinsu:retireRange) {
                if(tsi.getSalary()<=laborinsu) {
                    tsi.setRetireInsurance(laborinsu);
                    break;
                }
            }

            rtList.add(tsi);
        }

        return rtList;
    }

    public void insertBatch(final List<TempStaffInsurance> list){
        StringBuffer sql = new StringBuffer();
        sql.append(" Insert into TEMP_TEMPSTAFF_INSURANCE ");
        sql.append(" (INSURANCEYYMM, EMPNO, HOUR_RATE, HOURS, SALARY,LABOR_INSURANCE,HEALTH_INSURANCE, RETIRE_INSURANCE, ");
        sql.append(" UPDATE_DATE,UPDATE_USER,EMPNAME) ");
        sql.append(" values(?,?,?,?,?,?,?,?,?,?,?)");


        int[] count=DB.getJdbcTemplate().batchUpdate(sql.toString(), new BatchPreparedStatementSetter() {

            @Override
            public void setValues(PreparedStatement ps, int i) throws SQLException {
                TempStaffInsurance tsi = list.get(i);
                ps.setString(1, tsi.getInsuranceYyMm());
                ps.setString(2, tsi.getEmpNo());
                ps.setDouble(3, tsi.getHourRate());
                ps.setDouble(4, tsi.getHours());
                ps.setInt(5, tsi.getSalary());
                ps.setDouble(6, tsi.getLaborInsurance());
                ps.setDouble(7, tsi.getHealthInsurance());
                ps.setDouble(8, tsi.getRetireInsurance());
                ps.setTimestamp(9, new Timestamp(tsi.getUpdateDate().getTime()));
                ps.setString(10, tsi.getUpdateUser());
                ps.setString(11, tsi.getEmpName());
                System.out.println(tsi.toString());
            }

            @Override
            public int getBatchSize() {
                return list.size();
            }

        });
    }

    public void updateBatch(final List<TempStaffInsurance> list){
        StringBuffer sql = new StringBuffer();
        sql.append(" delete from TEMP_TEMPSTAFF_INSURANCE where INSURANCEYYMM=? and EMPNO=? ");

        int[] count=DB.getJdbcTemplate().batchUpdate(sql.toString(), new BatchPreparedStatementSetter() {

            @Override
            public void setValues(PreparedStatement ps, int i) throws SQLException {
                TempStaffInsurance tsi = list.get(i);
                ps.setString(1, tsi.getInsuranceYyMm());
                ps.setString(2, tsi.getEmpNo());
            }

            @Override
            public int getBatchSize() {
                return list.size();
            }

        });
    }

    public BigDecimal nonNullBigDecimal(Object obj) {
        if (obj == null) {
            return new BigDecimal("0");
        }
        return (BigDecimal) obj;
    }

}