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

import com.tymetro.ehrm.model.TempStaffSalary;
import com.tymetro.ehrm.utils.*;
import com.tymetro.ehrm.vo.TotalSalaryVO;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
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

import com.google.gson.JsonObject;

@Controller
@RequestMapping(path = "/temp/tempStaffController")
public class TempStaffController {

//    private final String UPLOADED_FOLDER = CFG.getWebAppRootPath()+File.separator+"upload";
    private final String UPLOADED_FOLDER = "/upload";
    private final String FILE_NAME = "tempExcel.xlsx";

    @Autowired
    ServletContext context;

    @ResponseBody
    @RequestMapping(path = "/tempstaff_salary_list")
    public String source(HttpServletRequest request) throws SQLException {
        String dateYM = request.getParameter("salaryYM");
        if (dateYM == null || dateYM.trim().equals("")) {
            dateYM = DB.queryScalar("select max(salaryYM) from temp_tempstaff_salary");
        }

        String sql = " select EMP_NO,SALARYYM, EMP_NAME, CAL_SALARY, LABOR_INSU, HEALTH_INSU, RETIRE_INSU from temp_tempstaff_salary ";
        sql += " where SALARYYM='" + dateYM + "' ";
        return DT.getDataTableRespond(request, sql);
    }

    @ResponseBody
    @RequestMapping(path = "/getRowData")
    public String getRowData(HttpServletRequest request) throws SQLException {
        JsonObject jsonResponse = new JsonObject();
        String empNo =request.getParameter("empNo");
        String salaryYM =request.getParameter("salaryYM");

        List<Map<String, Object>> empNoInsu = null;
        try {
            empNoInsu =DB.query("select * from TEMP_TEMPSTAFF_SALARY where EMP_NO=? and SALARYYM=?",new Object[] {empNo,salaryYM});
            jsonResponse.add("rowdata", JSON.toJson(empNoInsu));
        } catch(SQLException e) {
            jsonResponse.addProperty("message", e.toString());
        } catch (Exception e) {
            e.printStackTrace();
            jsonResponse.addProperty("message", "???????????????????????????????????????");
        }

        if(empNoInsu.size()==0) {
            jsonResponse.addProperty("message", "???????????????");
        }else {
            jsonResponse.addProperty("message", "????????????");
        }

        return jsonResponse.toString();
    }

    @ResponseBody
    @RequestMapping(path = "/insertRow")
    public String insertRow(HttpServletRequest request) throws SQLException {
        String userName=(String)request.getSession().getAttribute("tymetro_ehrm_user_name");
        TempStaffSalary tss = new TempStaffSalary();
        JsonObject jsonResponse = new JsonObject();

        try {
            tss.setSalaryYM((String) request.getParameter("salaryYM"));
            tss.setUnitName((String) request.getParameter("unitName"));
            tss.setEmpType((String) request.getParameter("empType"));
            tss.setEmpName((String) request.getParameter("empName"));
            tss.setEmpNo((String) request.getParameter("empNo"));
            tss.setSalary(Integer.parseInt(request.getParameter("salary")));
            tss.setHourRate(Double.parseDouble(request.getParameter("hourRate")));
            tss.setHours(Double.parseDouble(request.getParameter("hours")));
            tss.setDays(Double.parseDouble(request.getParameter("days")));
            tss.setRetirePay(Double.parseDouble( request.getParameter("retirePay"))/100);
            tss.setFamilyMem(Integer.parseInt(request.getParameter("familyMem")));
            tss.setHealthInsuMem(Double.parseDouble( request.getParameter("healthInsuMem")));
            tss.setLaborInsuMem(Double.parseDouble( request.getParameter("laborInsuMem")));
            tss.setOver65(request.getParameter("over65")==null?"N":(String)request.getParameter("over65"));
            tss.setLaborInsuDiff(Integer.parseInt(request.getParameter("laborInsuDiff")));
            tss.setLaborInsuComDiff(Integer.parseInt(request.getParameter("laborInsuComDiff")));
            tss.setHealthInsuDiff(Integer.parseInt(request.getParameter("healthInsuDiff")));
            tss.setHealthInsuComDiff(Integer.parseInt(request.getParameter("healthInsuComDiff")));
            tss.setRetireInsuDiff(Integer.parseInt(request.getParameter("retireInsuDiff")));
            tss.setRetireInsuComDiff(Integer.parseInt(request.getParameter("retireInsuComDiff")));
            tss.setOtherIn(Integer.parseInt(request.getParameter("otherIn")));
            tss.setOtherOut(Integer.parseInt(request.getParameter("otherOut")));
            tss.setOnboardDate((String) request.getParameter("onBoard"));
            tss.setResignationDate((String) request.getParameter("resignationDate"));
            tss.setLaborApplication(request.getParameter("laborApplication")==null?"N":(String)request.getParameter("laborApplication"));
            tss.setHealthApplication(request.getParameter("healthApplication")==null?"N":(String)request.getParameter("healthApplication"));
            tss.setWorkOvertime1(Double.parseDouble( request.getParameter("workOvertime1")));
            tss.setWorkOvertime2(Double.parseDouble( request.getParameter("workOvertime2")));
            tss.setWorkOvertime3(Double.parseDouble( request.getParameter("workOvertime3")));
            tss.setWorkOvertime4(Double.parseDouble( request.getParameter("workOvertime4")));
            tss.setWorkOvertime5(Double.parseDouble( request.getParameter("workOvertime5")));
            tss.setRemark((String) request.getParameter("remark"));
            List<Map<String, Object>> empNoInsu = null;
            try {
                empNoInsu =DB.query("select LABOR_INSURANCE,HEALTH_INSURANCE,RETIRE_INSURANCE from TEMP_TEMPSTAFF_INSURANCE where EMPNO=? and INSURANCEYYMM=?",new Object[] {tss.getEmpNo(),tss.getSalaryYM()});
            } catch (SQLException e) {
                e.printStackTrace();
                jsonResponse.addProperty("message", "????????????????????????????????????????????????");
                return jsonResponse.toString();
            }
            if(empNoInsu==null||empNoInsu.size()==0) {
                jsonResponse.addProperty("message", "????????????????????????????????????????????????");
                return jsonResponse.toString();
            }
            tss.setLaborInsu(((BigDecimal)empNoInsu.get(0).get("LABOR_INSURANCE")).intValue());
            tss.setHealthInsu(((BigDecimal)empNoInsu.get(0).get("HEALTH_INSURANCE")).intValue());
            tss.setRetireInsu(((BigDecimal)empNoInsu.get(0).get("RETIRE_INSURANCE")).intValue());

            tss.setUpdateDate(new Date());
            tss.setUpdateUser(userName);

            tss=calTempStaffSalary(tss);
            List<TempStaffSalary> list =new ArrayList<TempStaffSalary>();
            list.add(tss);
            updateBatch(list);
            insertBatch(list);

            jsonResponse.addProperty("message", "????????????????????????");
        } catch (Exception e) {
            e.printStackTrace();
            jsonResponse.addProperty("message", "?????????????????????"+e);
            return jsonResponse.toString();
        }

        return jsonResponse.toString();
    }


    @ResponseBody
    @RequestMapping(path = "/deleteRow")
    public String deleteRow(HttpServletRequest request) throws SQLException {
        String empNo = request.getParameter("empNo");
        String salaryYM = request.getParameter("salaryYM");
        JsonObject jsonResponse = new JsonObject();

        try {
            int rt=DB.getJdbcTemplate().update("delete from TEMP_TEMPSTAFF_SALARY where EMP_NO=? and salaryYM=?", new Object[] {empNo,salaryYM});
            jsonResponse.addProperty("message", "????????????"+rt+"?????????");
        } catch (Exception e) {
            e.printStackTrace();
            jsonResponse.addProperty("message", "?????????????????????"+e);
        }

        return jsonResponse.toString();
    }

//	@ResponseBody
//	@RequestMapping(path="/tempstaff_salary_list/show/{id}")
//	public String download(HttpServletResponse response, @PathVariable Integer id) throws SQLException {
//		String sql = "select log_info from log_db where log_id = ?";
//		String context = db.queryScalar(sql, new Object[]{id});
//
//		String name = "log(" + id + ")-" + DateUtil.getCurrentDate() + ".txt";
//		response.setContentType("text/plain");
//		response.setHeader("content-disposition", "attachment; filename=" + name);
//		response.setHeader("pragma", "no-cache");
//		response.setHeader("cache-control", "no-cache");
//		response.setDateHeader("expires", 0);
//		return context;
//	}

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

            List<TempStaffSalary> datalist = getFileList(UPLOADED_FOLDER+File.separator + FILE_NAME,userName);
            updateBatch(datalist);
            insertBatch(datalist);
            excelFile.delete();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            request.getSession().setAttribute("errorMsg", e.toString());
            return "redirect:/temp/temp_error.jsp";
        }
        return "redirect:/temp/temp_main.jsp";
    }

    @RequestMapping(value = "/download")
    public void download(HttpServletRequest request, HttpServletResponse response) throws SQLException {
        String salaryYM = request.getParameter("dateYM");
        XSSFWorkbook workbook;
        List<Map<String, Object>> list = DB.query("select * from TEMP_TEMPSTAFF_SALARY where SALARYYM='" + salaryYM + "' order by EMP_TYPE,EMP_NO ");
        List<Map<String, Object>> list2 = DB.query("select * from TEMP_TEMPSTAFF_SALARY where SALARYYM='" + salaryYM + "' order by EMP_TYPE,UNIT_NAME,EMP_NO ");
        try {
            File folder =new File(UPLOADED_FOLDER);
            if(!folder.exists()) {
                folder.mkdir();
            }
            String titleYear =(Integer.parseInt(salaryYM.substring(0,4))-1911)+"???";
            String titleMounth = salaryYM.substring(4,6)+"???";
            String fileName="????????????"+titleYear+titleMounth+".xlsx";
            fileName = URLEncoder.encode(fileName,"UTF-8");
            File file = new File(UPLOADED_FOLDER+File.separator + fileName);
            workbook = createExcel(new XSSFWorkbook(), list,list2, UPLOADED_FOLDER+File.separator + fileName);

            if (file.exists()) {
                String mimeType = context.getMimeType(file.getPath());

                if (mimeType == null) {
                    mimeType = "application/octet-stream";
                }

                response.setCharacterEncoding("UTF-8");
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
                System.out.println("Requested " + FILE_NAME + " file not found!!");
            }
            workbook.close();
        } catch (IOException e) {
            System.out.println("Error:- " + e.getMessage());
        }
    }

    @SuppressWarnings({ "rawtypes", "unchecked" })
    public List<TempStaffSalary> getFileList(String filePath,String userName) throws Exception {
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
                for (int i = 0; i <= 34; i++) {
                    if (currentRow.getCell(i) == null) {
                        list.add("");
                    } else if (currentRow.getCell(i).getCellTypeEnum() == CellType.STRING) {
                        list.add(currentRow.getCell(i).getStringCellValue());
                    } else if (currentRow.getCell(i).getCellTypeEnum() == CellType.NUMERIC) {
                        list.add(currentRow.getCell(i).getNumericCellValue());
                    }else if(currentRow.getCell(i).getCellTypeEnum()==CellType.BLANK){
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

        List<TempStaffSalary> rtList = new ArrayList();

        for (List list : excelList) {
            TempStaffSalary tss = new TempStaffSalary();
            try {
                tss.setSalaryYM((Integer.valueOf(((String) list.get(0)).substring(0, 3)) + 1911)
                        + ((String) list.get(0)).substring(4));// ????????????
                tss.setUnitName((String) list.get(1));// ????????????

                if ("?????????".equals((String) list.get(2))) {// ????????????
                    tss.setEmpType("S");
                } else if ("?????????".equals((String) list.get(2))) {
                    tss.setEmpType("M");
                } else if ("??????".equals((String) list.get(2))) {
                    tss.setEmpType("C");
                }

                tss.setEmpName((String) list.get(3));// ????????????
                tss.setEmpNo((String) list.get(4));// ????????????

                tss.setSalary("".equals(list.get(5))?0:((Double)list.get(5)).intValue());
                tss.setHourRate("".equals(list.get(6))?0:(Double)list.get(6));
                tss.setHours("".equals(list.get(7))?0:(Double)list.get(7));
                tss.setDays("".equals(list.get(8))?0:(Double)list.get(8));
                List<Map<String, Object>> empNoInsu = DB.query("select LABOR_INSURANCE,HEALTH_INSURANCE,RETIRE_INSURANCE from TEMP_TEMPSTAFF_INSURANCE where EMPNO=? and INSURANCEYYMM=?",new Object[] {tss.getEmpNo(),tss.getSalaryYM()});
                tss.setRetirePay(((Double) list.get(9)));// ????????????
                tss.setFamilyMem(((Double) list.get(10)).intValue());// ????????????
                tss.setHealthInsuMem((Double) list.get(11));// ??????????????????
                //tss.setLaborInsu(((Double) list.get(12)).intValue());// ????????????
                tss.setLaborInsu(((BigDecimal)empNoInsu.get(0).get("LABOR_INSURANCE")).intValue());
                tss.setLaborInsuMem((Double) list.get(13));// ??????????????????
                tss.setOver65((String)list.get(14));// ????????????65
                tss.setLaborInsuDiff(((Double) list.get(15)).intValue());// ??????????????????
                tss.setLaborInsuComDiff(((Double) list.get(16)).intValue());// ??????????????????
                //tss.setHealthInsu(((Double) list.get(17)).intValue());// ????????????
                tss.setHealthInsu(((BigDecimal)empNoInsu.get(0).get("HEALTH_INSURANCE")).intValue());
                tss.setHealthInsuDiff(((Double) list.get(18)).intValue());// ??????????????????
                tss.setHealthInsuComDiff(((Double) list.get(19)).intValue());// ??????????????????
                //tss.setRetireInsu(((Double) list.get(20)).intValue());// ????????????
                tss.setRetireInsu(((BigDecimal)empNoInsu.get(0).get("RETIRE_INSURANCE")).intValue());
                tss.setRetireInsuDiff(((Double) list.get(21)).intValue());// ??????????????????
                tss.setRetireInsuComDiff(((Double) list.get(22)).intValue());// ??????????????????
                tss.setOtherIn(((Double) list.get(23)).intValue());// ????????????
                tss.setOtherOut(((Double) list.get(24)).intValue());// ????????????
                tss.setOnboardDate((String)list.get(25));//?????????
                tss.setResignationDate((String)list.get(26));//?????????
                tss.setLaborApplication((String)list.get(27));//??????????????????
                tss.setHealthApplication((String)list.get(28));//??????????????????
                tss.setWorkOvertime1("".equals(list.get(29))==true?0:(Double)list.get(29));//????????????1.3334
                tss.setWorkOvertime2("".equals(list.get(30))==true?0:(Double)list.get(30));//????????????1.6667
                tss.setWorkOvertime3("".equals(list.get(31))==true?0:(Double)list.get(31));//????????????2.6667
                tss.setWorkOvertime4("".equals(list.get(32))==true?0:(Double)list.get(32));//????????????2
                tss.setWorkOvertime5("".equals(list.get(33))==true?0:(Double)list.get(33));//????????????1
                tss.setRemark("".equals(list.get(34))?"":list.get(34) instanceof Double?((Double)list.get(34)).toString():(String)list.get(34));//??????
            } catch (Exception e) {
                throw new Exception((String)list.get(3)+"  ????????????????????????????????????????????????????????????????????????????????????");
            }

            tss.setUpdateDate(new Date());
            tss.setUpdateUser(userName);

            /*
             * ??????????????????
             *
             */
            tss=calTempStaffSalary(tss);

            rtList.add(tss);
        }

        return rtList;
    }


    //???????????????????????????????????????CAL_XXXX??????????????????????????????
    @SuppressWarnings("deprecation")
    private TempStaffSalary calTempStaffSalary(TempStaffSalary tss) {

        // ?????????????????????
        if("N".equals(tss.getHealthApplication())) {
            tss.setHealthInsu(0);
        }
        if("N".equals(tss.getLaborApplication())) {
            tss.setLaborInsu(0);
        }
        //???????????? CAL_SALARY
        BigDecimal bd = BigDecimal.ZERO;
        if ("S".equals(tss.getEmpType())) {
            bd = new BigDecimal(tss.getHourRate() * tss.getHours());
        } else if ("M".equals(tss.getEmpType())) {
            bd = new BigDecimal(((double)tss.getHours() / 24) * tss.getSalary());
        } else if ("C".equals(tss.getEmpType())) {
            bd = new BigDecimal(tss.getDays() * tss.getSalary());
        }
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalSalary(bd.intValue());

        //?????????????????????????????????(???????????????)
        Double days=30.0;
        if(!tss.getOnboardDate().isEmpty()) {//?????????????????? 108/05/16
            days=30-Integer.parseInt(tss.getOnboardDate().substring(7))+1d;
            if(days==0) {//?????????31????????? ???????????????
                days=1.0;
            }
        }else if(!tss.getResignationDate().isEmpty()) {//??????????????????
            days=Double.parseDouble(tss.getResignationDate().substring(7));
        }

        bd = new BigDecimal(((tss.getLaborInsu() * 0.105 * 0.2* tss.getLaborInsuMem()/30.0*days)));
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        int i =bd.intValue();
        bd = new BigDecimal((tss.getLaborInsu() * 0.01 * 0.2* tss.getLaborInsuMem()/30.0*days));
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        int j=bd.intValue();

        tss.setCalLaborInsuSelf(i+j);// ??????????????????



        bd = new BigDecimal(((tss.getLaborInsu() * 0.1 * 0.7) + (tss.getLaborInsu() * 0.01 * 0.7))/30.0*days);
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalLaborInsuCom(bd.intValue());// ??????????????????

        //??????65???????????????????????????????????????????????????
        if("Y".equals(tss.getOver65())) {
            tss.setCalLaborInsuCom(0);// ??????????????????
        }

        bd = new BigDecimal(tss.getLaborInsu() * 0.0018/30.0*days);
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalLaborInsuComInjury(bd.intValue());// ??????????????????????????????

        if(!tss.getResignationDate().isEmpty()) {//???????????????????????????????????????
            bd = BigDecimal.ZERO;
        }else {
            bd = new BigDecimal(tss.getHealthInsu() * 0.0517 * 0.3 * tss.getHealthInsuMem());
        }
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalHealthInsuSelf(bd.intValue());// ??????????????????

        if(!tss.getResignationDate().isEmpty()) {//???????????????????????????????????????
            bd = BigDecimal.ZERO;
        }else {
            bd = new BigDecimal(tss.getHealthInsu() * 0.0517 * 0.6 * 1.58);
        }
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalHealthInsuCom(bd.intValue());// ??????????????????

        bd = new BigDecimal(tss.getRetireInsu() * tss.getRetirePay()/30.0*days);
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalRetireInsuSelf(bd.intValue());// ??????????????????

        bd = new BigDecimal(tss.getRetireInsu() * 0.06/30.0*days);
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalRetireInsuCom(bd.intValue());// ??????????????????

        bd = new BigDecimal(tss.getLaborInsu() * 0.00025/30.0*days);
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalRepaymentFund(bd.intValue());// ????????????

        if ("S".equals(tss.getEmpType())) {// ?????????
            bd = new BigDecimal(tss.getCalSalary() * 0.005);
        } else if ("M".equals(tss.getEmpType())) {
            bd = new BigDecimal(tss.getCalSalary() * 0.005);
        } else if ("C".equals(tss.getEmpType())) {
            bd =BigDecimal.ZERO;//????????????????????????NO REASON
        }
        bd = bd.setScale(0, BigDecimal.ROUND_HALF_UP);
        tss.setCalWelfare(bd.intValue());

        //???????????????
        bd = new BigDecimal(tss.getHourRate()*(
                (tss.getWorkOvertime1()*1.3334)+
                        (tss.getWorkOvertime2()*1.6667)+
                        (tss.getWorkOvertime3()*2.6667)+
                        (tss.getWorkOvertime4()*2)+
                        (tss.getWorkOvertime5()*1)));
        bd = bd.setScale(0, BigDecimal.ROUND_CEILING);
        tss.setCalSalary(bd.intValue()+tss.getCalSalary());
        return tss;
    }


    public void insertBatch(final List<TempStaffSalary> list){
        StringBuffer sql = new StringBuffer();
        sql.append(" Insert into TEMP_TEMPSTAFF_SALARY ");
        sql.append(" (SALARYYM, UNIT_NAME, EMP_TYPE, EMP_NAME, EMP_NO,SALARY,HOUR_RATE, HOURS, DAYS, RETIRE_PAY, FAMILY_MEM, ");
        sql.append(" HEALTH_INSU_MEM, LABOR_INSU, LABOR_INSU_MEM, OVER_65, LABOR_INSU_DIFF,LABOR_INSU_COM_DIFF, HEALTH_INSU, ");
        sql.append(" HEALTH_INSU_DIFF, HEALTH_INSU_COM_DIFF, RETIRE_INSU,RETIRE_INSU_DIFF, RETIRE_INSU_COM_DIFF, OTHER_IN, ");
        sql.append(" OTHER_OUT,CAL_LABOR_INSU_SELF,CAL_LABOR_INSU_COM,CAL_LABOR_INSU_COM_INJURY,CAL_RETIRE_INSU_SELF, ");
        sql.append(" CAL_RETIRE_INSU_COM,CAL_HEALTH_INSU_SELF,CAL_HEALTH_INSU_COM,CAL_REPAYMENT_FUND,CAL_WELFARE, ");
        sql.append(" ONBOARD_DATE,RESIGNATION_DATE,UPDATE_DATE,UPDATE_USER,LABOR_APPLICATION,HEALTH_APPLICATION, ");
        sql.append(" WORK_OVERTIME1,WORK_OVERTIME2,WORK_OVERTIME3,WORK_OVERTIME4,WORK_OVERTIME5,CAL_SALARY,REMARK) ");
        sql.append(" values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");


        int[] count=DB.getJdbcTemplate().batchUpdate(sql.toString(), new BatchPreparedStatementSetter() {

            @Override
            public void setValues(PreparedStatement ps, int i) throws SQLException {
                TempStaffSalary tempStaffSalary = list.get(i);
                ps.setString(1, tempStaffSalary.getSalaryYM());
                ps.setString(2, tempStaffSalary.getUnitName());
                ps.setString(3, tempStaffSalary.getEmpType());
                ps.setString(4, tempStaffSalary.getEmpName());
                ps.setString(5, tempStaffSalary.getEmpNo());
                ps.setDouble(6, tempStaffSalary.getSalary());
                ps.setDouble(7, tempStaffSalary.getHourRate());
                ps.setDouble(8, tempStaffSalary.getHours());
                ps.setDouble(9, tempStaffSalary.getDays());
                ps.setInt(10, tempStaffSalary.getRetirePay().intValue());
                ps.setInt(11, tempStaffSalary.getFamilyMem().intValue());
                ps.setInt(12, tempStaffSalary.getHealthInsuMem().intValue());
                ps.setInt(13, tempStaffSalary.getLaborInsu().intValue());
                ps.setInt(14, tempStaffSalary.getLaborInsuMem().intValue());
                ps.setString(15, tempStaffSalary.getOver65());
                ps.setInt(16, tempStaffSalary.getLaborInsuDiff().intValue());
                ps.setInt(17, tempStaffSalary.getLaborInsuComDiff().intValue());
                ps.setInt(18, tempStaffSalary.getHealthInsu().intValue());
                ps.setInt(19, tempStaffSalary.getHealthInsuDiff().intValue());
                ps.setInt(20, tempStaffSalary.getHealthInsuComDiff().intValue());
                ps.setInt(21, tempStaffSalary.getRetireInsu().intValue());
                ps.setInt(22, tempStaffSalary.getRetireInsuDiff().intValue());
                ps.setInt(23, tempStaffSalary.getRetireInsuComDiff().intValue());
                ps.setInt(24, tempStaffSalary.getOtherIn().intValue());
                ps.setInt(25, tempStaffSalary.getOtherOut().intValue());
                ps.setInt(26, tempStaffSalary.getCalLaborInsuSelf().intValue());
                ps.setInt(27, tempStaffSalary.getCalLaborInsuCom().intValue());
                ps.setInt(28, tempStaffSalary.getCalLaborInsuComInjury().intValue());
                ps.setInt(29, tempStaffSalary.getCalRetireInsuSelf().intValue());
                ps.setInt(30, tempStaffSalary.getCalRetireInsuCom().intValue());
                ps.setInt(31, tempStaffSalary.getCalHealthInsuSelf().intValue());
                ps.setInt(32, tempStaffSalary.getCalHealthInsuCom().intValue());
                ps.setInt(33, tempStaffSalary.getCalRepaymentFund().intValue());
                ps.setInt(34, tempStaffSalary.getCalWelfare().intValue());
                ps.setString(35, tempStaffSalary.getOnboardDate());
                ps.setString(36, tempStaffSalary.getResignationDate());
                ps.setTimestamp(37, new Timestamp(tempStaffSalary.getUpdateDate().getTime()));
                ps.setString(38, tempStaffSalary.getUpdateUser());
                ps.setString(39, tempStaffSalary.getLaborApplication());
                ps.setString(40, tempStaffSalary.getHealthApplication());
                ps.setDouble(41, tempStaffSalary.getWorkOvertime1());
                ps.setDouble(42, tempStaffSalary.getWorkOvertime2());
                ps.setDouble(43, tempStaffSalary.getWorkOvertime3());
                ps.setDouble(44, tempStaffSalary.getWorkOvertime4());
                ps.setDouble(45, tempStaffSalary.getWorkOvertime5());
                ps.setDouble(46, tempStaffSalary.getCalSalary());
                ps.setString(47, tempStaffSalary.getRemark());
            }

            @Override
            public int getBatchSize() {
                return list.size();
            }

        });
    }


    public void updateBatch(final List<TempStaffSalary> list){
        StringBuffer sql = new StringBuffer();
        sql.append(" delete from TEMP_TEMPSTAFF_SALARY where SALARYYM=? and EMP_NO=? ");

        int[] count=DB.getJdbcTemplate().batchUpdate(sql.toString(), new BatchPreparedStatementSetter() {

            @Override
            public void setValues(PreparedStatement ps, int i) throws SQLException {
                TempStaffSalary tempStaffSalary = list.get(i);
                ps.setString(1, tempStaffSalary.getSalaryYM());
                ps.setString(2, tempStaffSalary.getEmpNo());
            }

            @Override
            public int getBatchSize() {
                return list.size();
            }

        });
    }


    public XSSFWorkbook createExcel(XSSFWorkbook workbook, List<Map<String, Object>> list,List<Map<String, Object>> list2, String path) {

        CellStyle style1 = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        XSSFDataFormat format= workbook.createDataFormat();
        fontStyle1.setFontName("?????????");
        fontStyle1.setFontHeightInPoints((short) 10);
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setFont(fontStyle1);
        style1.setDataFormat(format.getFormat("###,##0"));

        String a = "";
        String b = "";
        int rows = 0;
        XSSFSheet sheet = null;
        TotalSalaryVO totalSalaryVO =new TotalSalaryVO();//EMP_TYPE ??????
        totalSalaryVO.initValue();//?????????
        TotalSalaryVO totalSalaryVO2 =new TotalSalaryVO();//EMP_TYPE +UNIT_NAME ??????
        totalSalaryVO2.initValue();//?????????
        // 26??????
        int[] cellWidth = { 2000, 2000, 2400, 2400, 2400, 2400, 2400, 2000, 2000, 2000, 2400, 2400, 2400, 2400,
                2400, 2000, 2400, 2600, 2600, 2000, 2000, 2000, 2000, 2000, 2000, 2400 };

        for (Map map : list) {
            b = (String) map.get("EMP_TYPE");
            if (a.equals("")) {
                a = (String) map.get("EMP_TYPE");
                sheet = workbook.createSheet(getSheetName(a));
                for (int i = 0; i < cellWidth.length; i++) {
                    sheet.setColumnWidth(i, cellWidth[i]);
                }
                rows = setSheetHeader(workbook, sheet, rows, cellWidth.length,(String)map.get("SALARYYM"));

            } else if (a.equals(b)) {

            } else {
                rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO);//????????????
                rows=setSheetLineComment(workbook,sheet, style1, rows);
                totalSalaryVO.initValue();
                sheet = workbook.createSheet(getSheetName(b));
                for (int i = 0; i < cellWidth.length; i++) {
                    sheet.setColumnWidth(i, cellWidth[i]);
                }
                rows = 0;
                rows = setSheetHeader(workbook, sheet, rows, cellWidth.length,(String)map.get("SALARYYM"));

                a = b;
            }

            rows = setSheetLine(workbook,sheet, style1, rows, map,totalSalaryVO,totalSalaryVO2);//???????????????????????????

            sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);//??????A4??????
            sheet.getPrintSetup().setLandscape(true);//????????????
            sheet.setFitToPage(true);//???????????????????????????
            sheet.setSelected(true);
        }
        rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO);//????????????
        rows=setSheetLineComment(workbook,sheet, style1, rows);
        /*
         * ??????????????????
         */

        a = "";
        b = "";
        rows = 0;
        totalSalaryVO.initValue();//?????????
        totalSalaryVO2.initValue();//?????????
        String c="";
        String d="";
        for (Map map : list2) {
            b = (String) map.get("EMP_TYPE");
            if (a.equals("")) {
                a = (String) map.get("EMP_TYPE");
                sheet = workbook.createSheet(getSheetName(a)+"BY ??????");
                for (int i = 0; i < cellWidth.length; i++) {
                    sheet.setColumnWidth(i, cellWidth[i]);
                }
                rows = setSheetHeader(workbook, sheet, rows, cellWidth.length,(String)map.get("SALARYYM"));

            } else if (a.equals(b)) {

            } else {
                rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO2);//????????????
                rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO);//????????????
                rows=setSheetLineComment(workbook,sheet, style1, rows);
                totalSalaryVO.initValue();
                sheet = workbook.createSheet(getSheetName(b)+"BY ??????");
                for (int i = 0; i < cellWidth.length; i++) {
                    sheet.setColumnWidth(i, cellWidth[i]);
                }
                rows = 0;
                rows = setSheetHeader(workbook, sheet, rows, cellWidth.length,(String)map.get("SALARYYM"));

                a = b;

                totalSalaryVO2.initValue();
                c="";
                d="";
            }

            d= (String) map.get("UNIT_NAME");
            if("".equals(c)) {
                c= (String) map.get("UNIT_NAME");
            }else if(c.equals(d)) {

            }else {
                rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO2);//????????????
                totalSalaryVO2.initValue();
                c= d;
            }

            rows = setSheetLine(workbook,sheet, style1, rows, map,totalSalaryVO,totalSalaryVO2);

            sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);//??????A4??????
            sheet.getPrintSetup().setLandscape(true);//????????????
            sheet.setFitToPage(true);//???????????????????????????
            sheet.setSelected(true);
        }
        rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO2);//????????????
        rows=setSheetLineTotal(workbook,sheet, style1, rows,totalSalaryVO);//????????????
        rows=setSheetLineComment(workbook,sheet, style1, rows);

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



    public String getSheetName(String s) {
        if (s.equals("S")) {
            return "?????????";
        } else if (s.equals("M")) {
            return "?????????";
        } else if (s.equals("C")) {
            return "??????";
        } else {
            return "UNKNOW";
        }
    }


    public int setSheetHeader(XSSFWorkbook workbook, XSSFSheet sheet, int rows, int length,String titleYM ) {
        XSSFRow xssfRow;
        XSSFCell cell;
        String titleYear =(Integer.parseInt(titleYM.substring(0,4))-1911)+"???";
        String titleMounth = titleYM.substring(4,6)+"???";
        CellStyle style1 = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        fontStyle1.setBold(true);
        fontStyle1.setFontName("?????????");
        fontStyle1.setFontHeightInPoints((short) 18);
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setFont(fontStyle1);

        CellStyle style2 = workbook.createCellStyle();
        Font fontStyle2 = workbook.createFont();
        fontStyle2.setBold(true);
        fontStyle2.setFontName("?????????");
        fontStyle2.setFontHeightInPoints((short) 10);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setFont(fontStyle2);
        style2.setWrapText(true);

        CellStyle style3 = workbook.createCellStyle();
        Font fontStyle3 = workbook.createFont();
        fontStyle3.setFontName("?????????");
        fontStyle3.setFontHeightInPoints((short) 10);
        style3.setAlignment(HorizontalAlignment.CENTER);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);
        style3.setBorderBottom(BorderStyle.THIN);
        style3.setBorderLeft(BorderStyle.THIN);
        style3.setBorderRight(BorderStyle.THIN);
        style3.setBorderTop(BorderStyle.THIN);
        style3.setFont(fontStyle3);

        CellStyle styleBR = workbook.createCellStyle();
        styleBR.setBorderBottom(BorderStyle.THIN);
        styleBR.setBorderRight(BorderStyle.THIN);

        CellStyle styleB = workbook.createCellStyle();
        styleB.setBorderBottom(BorderStyle.THIN);

        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 600);
        cell = xssfRow.createCell(0);
        cell.setCellValue("????????????????????????????????????"+titleYear+titleMounth+"????????????");
        cell.setCellStyle(style1);
        for (int i = 1; i < length; i++) {
            cell = xssfRow.createCell(i);
            cell.setCellStyle(style1);
        }
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 0, length - 1));

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        cell = xssfRow.createCell(0);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 0, 1));

        cell = xssfRow.createCell(2);
        cell.setCellValue("");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 2, 4));

        cell = xssfRow.createCell(5);
        cell.setCellValue("?????????????????????????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 5, 15));

        cell = xssfRow.createCell(8);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(16);
        cell.setCellValue("");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(17);
        cell.setCellValue("???????????????????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 17, 18));

        cell = xssfRow.createCell(18);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(19);
        cell.setCellValue("??????????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 19, 22));

        cell = xssfRow.createCell(23);
        cell.setCellValue("??????????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 23, 24));

        cell = xssfRow.createCell(25);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 25, 25));

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        cell = xssfRow.createCell(0);
        cell.setCellValue("??????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(1);
        cell.setCellValue("??????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(2);
        cell.setCellValue("???????????? ??????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 2, 2));

        cell = xssfRow.createCell(3);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 3, 3));

        cell = xssfRow.createCell(4);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 4, 4));

        cell = xssfRow.createCell(5);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(6);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 6, 6));

        cell = xssfRow.createCell(7);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 7, 8));

        cell = xssfRow.createCell(9);
        cell.setCellValue("?????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(10);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 10, 10));

        cell = xssfRow.createCell(11);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 11, 11));

        cell = xssfRow.createCell(12);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);

        cell = xssfRow.createCell(13);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 13, 13));

        cell = xssfRow.createCell(14);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 14, 14));

        cell = xssfRow.createCell(15);
        cell.setCellValue("??????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 15, 15));

        cell = xssfRow.createCell(16);
        cell.setCellValue("????????????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows + 1, 16, 16));

        cell = xssfRow.createCell(17);
        cell.setCellValue("??????");
        cell.setCellStyle(style2);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 17, 18));

        cell = xssfRow.createCell(18);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(22);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(24);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(25);
        cell.setCellStyle(styleBR);

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        cell = xssfRow.createCell(0);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(1);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(2);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(3);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(4);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(5);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(6);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(7);
        cell.setCellValue("????????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(8);
        cell.setCellValue("????????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(9);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(10);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(11);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(12);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(13);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(14);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(15);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(16);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(17);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(18);
        cell.setCellValue("????????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(19);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(20);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(21);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(22);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(23);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(24);
        cell.setCellValue("?????????");
        cell.setCellStyle(style3);

        cell = xssfRow.createCell(25);
        cell.setCellValue("??????");
        cell.setCellStyle(style3);

        return rows;
    }


    @SuppressWarnings("deprecation")
    public int setSheetLine(XSSFWorkbook workbook,XSSFSheet sheet, CellStyle cellStyle, int rows, Map map,TotalSalaryVO totalSalaryVO,TotalSalaryVO totalSalaryVO2) {
        XSSFRow xssfRow;
        XSSFCell cell;
        CellStyle style1 = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        XSSFDataFormat format= workbook.createDataFormat();
        fontStyle1.setBold(true);
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);
        style1.setFont(fontStyle1);
        style1.setDataFormat(format.getFormat("###,##0"));

        totalSalaryVO.setCountNO(totalSalaryVO.getCountNO()+1);//EXCEL????????????
        totalSalaryVO2.setCountNO(totalSalaryVO2.getCountNO()+1);//EXCEL????????????

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 400);
        cell = xssfRow.createCell(0);// ??????
        cell.setCellValue((String) map.get("EMP_NAME"));
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(1);// ??????
        cell.setCellValue((String) map.get("UNIT_NAME"));
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(2);// ????????????or??????
        if("S".equals((String)map.get("EMP_TYPE"))) {
            cell.setCellValue(((BigDecimal)map.get("HOURS")).toString());
        }else {
            cell.setCellValue(nonNullBigDecimal(map.get("CAL_SALARY")).doubleValue());
        }
        cell.setCellStyle(cellStyle);

        BigDecimal workOverTime =nonNullBigDecimal(map.get("HOUR_RATE")).multiply(
                nonNullBigDecimal(map.get("WORK_OVERTIME1")).multiply(new BigDecimal("1.3334")).add(
                        nonNullBigDecimal(map.get("WORK_OVERTIME2")).multiply(new BigDecimal("1.6667"))).add(
                        nonNullBigDecimal(map.get("WORK_OVERTIME3")).multiply(new BigDecimal("2.6667"))).add(
                        nonNullBigDecimal(map.get("WORK_OVERTIME4")).multiply(new BigDecimal("2"))).add(
                        nonNullBigDecimal(map.get("WORK_OVERTIME5")).multiply(new BigDecimal("1")))
        );

        cell = xssfRow.createCell(3);// ????????????
//		workOverTime = workOverTime.setScale(0, BigDecimal.ROUND_HALF_UP);
        cell.setCellValue(((BigDecimal) map.get("OTHER_IN")).add(workOverTime.setScale(0, BigDecimal.ROUND_CEILING)).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(4);// ????????????
        BigDecimal salary = ((BigDecimal) map.get("CAL_SALARY")).add((BigDecimal) map.get("OTHER_IN"));

        totalSalaryVO.setSalaryTotal(totalSalaryVO.getSalaryTotal().add(salary));//EXCEL??????????????????
        totalSalaryVO2.setSalaryTotal(totalSalaryVO2.getSalaryTotal().add(salary));//EXCEL??????????????????
        cell.setCellValue(salary.doubleValue());
        cell.setCellStyle(style1);

        cell = xssfRow.createCell(5);// ????????????
        cell.setCellValue(((BigDecimal) map.get("LABOR_INSU")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(6);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_LABOR_INSU_SELF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(7);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_LABOR_INSU_COM")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(8);// ????????????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_LABOR_INSU_COM_INJURY")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(9);// ????????????
        cell.setCellValue(((BigDecimal) map.get("RETIRE_INSU")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(10);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_RETIRE_INSU_SELF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(11);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_RETIRE_INSU_COM")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(12);// ????????????
        cell.setCellValue(((BigDecimal) map.get("HEALTH_INSU")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(13);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_HEALTH_INSU_SELF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(14);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_HEALTH_INSU_COM")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(15);// ????????????
        cell.setCellValue(((BigDecimal) map.get("FAMILY_MEM")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(16);// ????????????
        cell.setCellValue(((BigDecimal) map.get("CAL_REPAYMENT_FUND")).doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setRepaymentFundTotal(totalSalaryVO.getRepaymentFundTotal().add((BigDecimal) map.get("CAL_REPAYMENT_FUND")));//EXCEL??????????????????
        totalSalaryVO2.setRepaymentFundTotal(totalSalaryVO2.getRepaymentFundTotal().add((BigDecimal) map.get("CAL_REPAYMENT_FUND")));//EXCEL??????????????????

        cell = xssfRow.createCell(17);// ?????????
        cell.setCellValue(((BigDecimal) map.get("CAL_WELFARE")).doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setWelfare(totalSalaryVO.getWelfare().add((BigDecimal) map.get("CAL_WELFARE")));//EXCEL???????????????
        totalSalaryVO2.setWelfare(totalSalaryVO2.getWelfare().add((BigDecimal) map.get("CAL_WELFARE")));//EXCEL???????????????

        cell = xssfRow.createCell(18);// ????????????
        cell.setCellValue(((BigDecimal) map.get("OTHER_OUT")).doubleValue());
        cell.setCellStyle(cellStyle);


        BigDecimal laborSelf =((BigDecimal) map.get("CAL_LABOR_INSU_SELF")).add((BigDecimal) map.get("LABOR_INSU_DIFF"));
        BigDecimal laborCom =((BigDecimal) map.get("CAL_LABOR_INSU_COM")).add((BigDecimal) map.get("LABOR_INSU_COM_DIFF"));
        BigDecimal retireSelf =((BigDecimal) map.get("CAL_RETIRE_INSU_SELF")).add((BigDecimal) map.get("RETIRE_INSU_DIFF"));
        BigDecimal retireCom =((BigDecimal) map.get("CAL_RETIRE_INSU_COM")).add((BigDecimal) map.get("RETIRE_INSU_COM_DIFF"));
        BigDecimal healthSelf =((BigDecimal) map.get("CAL_HEALTH_INSU_SELF")).add((BigDecimal) map.get("HEALTH_INSU_DIFF"));
        BigDecimal healthCom =((BigDecimal) map.get("CAL_HEALTH_INSU_COM")).add((BigDecimal) map.get("HEALTH_INSU_COM_DIFF"));
        BigDecimal insuTotal = (laborSelf.add(retireSelf).add(healthSelf));

        cell = xssfRow.createCell(19);// ??????
        cell.setCellValue(insuTotal.doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(20);// ???????????????
        cell.setCellValue(laborSelf.doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setLaborSelfTotal(totalSalaryVO.getLaborSelfTotal().add(laborSelf));//EXCEL??????????????????
        totalSalaryVO2.setLaborSelfTotal(totalSalaryVO2.getLaborSelfTotal().add(laborSelf));//EXCEL??????????????????

        cell = xssfRow.createCell(21);// ???????????????
        cell.setCellValue(healthSelf.doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setHealthSelfTotal(totalSalaryVO.getHealthSelfTotal().add(healthSelf));//EXCEL??????????????????
        totalSalaryVO2.setHealthSelfTotal(totalSalaryVO2.getHealthSelfTotal().add(healthSelf));//EXCEL??????????????????

        cell = xssfRow.createCell(22);// ???????????????
        cell.setCellValue(retireSelf.doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setRetireSelfTotal(totalSalaryVO.getRetireSelfTotal().add(retireSelf));//EXCEL??????????????????
        totalSalaryVO2.setRetireSelfTotal(totalSalaryVO2.getRetireSelfTotal().add(retireSelf));//EXCEL??????????????????

        cell = xssfRow.createCell(23);// ???????????????
        cell.setCellValue(laborCom.add((BigDecimal) map.get("CAL_LABOR_INSU_COM_INJURY")).add(healthCom).doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setLaborComTotal(totalSalaryVO.getLaborComTotal().add(laborCom).add((BigDecimal) map.get("CAL_LABOR_INSU_COM_INJURY")));//EXCEL??????????????????
        totalSalaryVO.setHealthComTotal(totalSalaryVO.getHealthComTotal().add(healthCom));//EXCEL??????????????????
        totalSalaryVO2.setLaborComTotal(totalSalaryVO2.getLaborComTotal().add(laborCom).add((BigDecimal) map.get("CAL_LABOR_INSU_COM_INJURY")));//EXCEL??????????????????
        totalSalaryVO2.setHealthComTotal(totalSalaryVO2.getHealthComTotal().add(healthCom));//EXCEL??????????????????

        cell = xssfRow.createCell(24);// ????????????
        cell.setCellValue(retireCom.doubleValue());
        cell.setCellStyle(cellStyle);
        totalSalaryVO.setRetireComTotal(totalSalaryVO.getRetireComTotal().add(retireCom));//EXCEL??????????????????
        totalSalaryVO2.setRetireComTotal(totalSalaryVO2.getRetireComTotal().add(retireCom));//EXCEL??????????????????

        cell = xssfRow.createCell(25);
        BigDecimal result = salary.subtract(insuTotal.add((BigDecimal) map.get("CAL_WELFARE"))
                .add((BigDecimal) map.get("OTHER_OUT")));
        cell.setCellValue(result.doubleValue());
        cell.setCellStyle(style1);
        totalSalaryVO.setSalaryFinalTotal(totalSalaryVO.getSalaryFinalTotal().add(result));//EXCEL??????????????????
        totalSalaryVO2.setSalaryFinalTotal(totalSalaryVO2.getSalaryFinalTotal().add(result));//EXCEL??????????????????

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 400);
        cell = xssfRow.createCell(0);// ??????
        cell.setCellValue((String) map.get("EMP_NO"));
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(1);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(2);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(3);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(4);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(5);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(6);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("LABOR_INSU_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(7);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(8);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("LABOR_INSU_COM_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(9);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(10);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("RETIRE_INSU_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(11);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("RETIRE_INSU_COM_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(12);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(13);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("HEALTH_INSU_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(14);// ??????????????????
        cell.setCellValue(((BigDecimal) map.get("HEALTH_INSU_COM_DIFF")).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(15);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(16);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(17);
        String s1="";
        if(!TEXT.isEmpty((String) map.get("ONBOARD_DATE"))||!TEXT.isEmpty((String) map.get("RESIGNATION_DATE"))) {
            if(!TEXT.isEmpty((String) map.get("ONBOARD_DATE")))
                s1+=(String) map.get("ONBOARD_DATE")+"??????";
            if(!TEXT.isEmpty((String) map.get("RESIGNATION_DATE")))
                s1+=(String) map.get("RESIGNATION_DATE")+"??????";
        }
        s1+=TEXT.isEmpty((String) map.get("REMARK"))?"":(String) map.get("REMARK");
        cell.setCellValue(s1);
        cell.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 17, 18));

        cell = xssfRow.createCell(18);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(19);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(20);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(21);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(22);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(23);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(24);
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(25);
        cell.setCellStyle(cellStyle);

        return rows;
    }

    public int setSheetLineTotal(XSSFWorkbook workbook,XSSFSheet sheet, CellStyle cellStyle, int rows,TotalSalaryVO totalSalaryVO) {
        XSSFRow xssfRow;
        XSSFCell cell;
        XSSFDataFormat format= workbook.createDataFormat();

        CellStyle styleLeft = workbook.createCellStyle();
        Font fontStyle1 = workbook.createFont();
        fontStyle1.setFontName("?????????");
        fontStyle1.setFontHeightInPoints((short) 10);
        styleLeft.setAlignment(HorizontalAlignment.LEFT);
        styleLeft.setVerticalAlignment(VerticalAlignment.CENTER);
        styleLeft.setBorderBottom(BorderStyle.THIN);
        styleLeft.setBorderLeft(BorderStyle.THIN);
        styleLeft.setBorderRight(BorderStyle.THIN);
        styleLeft.setBorderTop(BorderStyle.THIN);
        styleLeft.setFont(fontStyle1);

        CellStyle styleBR = workbook.createCellStyle();
        styleBR.setBorderBottom(BorderStyle.THIN);
        styleBR.setBorderRight(BorderStyle.THIN);

        CellStyle styleB = workbook.createCellStyle();
        styleB.setBorderBottom(BorderStyle.THIN);

        CellStyle styleBold = workbook.createCellStyle();
        Font fontStyleBold = workbook.createFont();
        fontStyleBold.setBold(true);
        styleBold.setAlignment(HorizontalAlignment.CENTER);
        styleBold.setVerticalAlignment(VerticalAlignment.CENTER);
        styleBold.setBorderBottom(BorderStyle.THIN);
        styleBold.setBorderLeft(BorderStyle.THIN);
        styleBold.setBorderRight(BorderStyle.THIN);
        styleBold.setBorderTop(BorderStyle.THIN);
        styleBold.setFont(fontStyleBold);
        styleBold.setDataFormat(format.getFormat("###,##0"));

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 400);
        cell = xssfRow.createCell(0);
        cell.setCellValue("???"+totalSalaryVO.getCountNO()+"???");//?????????
        cell.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 0, 2));

        cell = xssfRow.createCell(1);
        cell.setCellStyle(styleB);
        cell = xssfRow.createCell(2);
        cell.setCellStyle(styleB);
        cell = xssfRow.createCell(3);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(4);
        cell.setCellValue(totalSalaryVO.getSalaryTotal().doubleValue());//????????????
        cell.setCellStyle(styleBold);

        cell = xssfRow.createCell(5);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(6);
        cell.setCellValue(totalSalaryVO.getLaborSelfTotal().doubleValue());//????????????
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(7);
        cell.setCellValue(totalSalaryVO.getLaborComTotal().doubleValue());//??????????????????
        cell.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 7, 8));

        cell = xssfRow.createCell(8);
        cell.setCellStyle(styleBR);

        cell = xssfRow.createCell(9);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(10);
        cell.setCellValue(totalSalaryVO.getRetireSelfTotal().doubleValue());//????????????
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(11);
        cell.setCellValue(totalSalaryVO.getRetireComTotal().doubleValue());//????????????
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(12);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(13);
        cell.setCellValue(totalSalaryVO.getHealthSelfTotal().doubleValue());//????????????
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(14);
        cell.setCellValue(totalSalaryVO.getHealthComTotal().doubleValue());//????????????
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(15);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(16);//????????????
        cell.setCellValue(totalSalaryVO.getRepaymentFundTotal().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(17);//?????????
        cell.setCellValue(totalSalaryVO.getWelfare().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(18);
        cell.setCellStyle(styleB);

        cell = xssfRow.createCell(19);//??????
        cell.setCellValue(totalSalaryVO.getLaborSelfTotal().add(totalSalaryVO.getHealthSelfTotal()).add(totalSalaryVO.getRetireSelfTotal()).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(20);//??????
        cell.setCellValue(totalSalaryVO.getLaborSelfTotal().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(21);//??????
        cell.setCellValue(totalSalaryVO.getHealthSelfTotal().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(22);//??????
        cell.setCellValue(totalSalaryVO.getRetireSelfTotal().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(23);//???????????????
        cell.setCellValue(totalSalaryVO.getLaborComTotal().add(totalSalaryVO.getHealthComTotal()).doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(24);//????????????
        cell.setCellValue(totalSalaryVO.getRetireComTotal().doubleValue());
        cell.setCellStyle(cellStyle);

        cell = xssfRow.createCell(25);//????????????
        cell.setCellValue(totalSalaryVO.getSalaryFinalTotal().doubleValue());
        cell.setCellStyle(styleBold);

        return rows;
    }

    public int setSheetLineComment(XSSFWorkbook workbook,XSSFSheet sheet, CellStyle cellStyle, int rows) {
        XSSFRow xssfRow;
        XSSFCell cell;

        rows = rows + 1;
        xssfRow = sheet.createRow(rows);
        xssfRow.setHeight((short) 400);
        cell = xssfRow.createCell(0);
        cell.setCellValue("?????? ??? ");

        cell = xssfRow.createCell(5);
        cell.setCellValue("???????????? ");

        cell = xssfRow.createCell(10);
        cell.setCellValue("???????????? ??? ");

        cell = xssfRow.createCell(15);
        cell.setCellValue("?????????????????????????????????");

        return rows;
    }

    public BigDecimal nonNullBigDecimal(Object obj) {
        if (obj == null) {
            return new BigDecimal("0");
        }
        return (BigDecimal) obj;
    }

}