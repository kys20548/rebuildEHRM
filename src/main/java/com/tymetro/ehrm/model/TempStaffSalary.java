package com.tymetro.ehrm.model;

import lombok.*;

import javax.persistence.*;
import java.util.Date;

@Data
@Entity(name = "TEMP_STAFF_SALARY")
public class TempStaffSalary {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private int seq;
    private String salaryYM;
    private String unitName;
    private String empType;
    private String empName;
    private String empNo;
    private Integer salary;
    private Double hourRate;
    private Double hours;
    private Double days;
    private Double retirePay;
    private Integer familyMem;
    private Double healthInsuMem;
    private Integer laborInsu;
    private Double laborInsuMem;
    private String over65;
    private Integer laborInsuDiff;
    private Integer laborInsuComDiff;
    private Integer healthInsu;
    private Integer healthInsuDiff;
    private Integer healthInsuComDiff;
    private Integer retireInsu;
    private Integer retireInsuDiff;
    private Integer retireInsuComDiff;
    private Integer otherIn;
    private Integer otherOut;
    private String onboardDate;
    private String resignationDate;
    private String laborApplication;
    private String healthApplication;
    private Double workOvertime1;
    private Double workOvertime2;
    private Double workOvertime3;
    private Double workOvertime4;
    private Double workOvertime5;
    private String remark;
    private Integer calSalary;
    private Integer calLaborInsuSelf;
    private Integer calLaborInsuCom;
    private Integer calLaborInsuComInjury;
    private Integer calHealthInsuSelf;
    private Integer calHealthInsuCom;
    private Integer calRetireInsuSelf;
    private Integer calRetireInsuCom;
    private Integer calRepaymentFund;
    private Integer calWelfare;
    private Date updateDate;
    private String updateUser;

}
