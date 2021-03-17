package com.tymetro.ehrm.model;

import lombok.*;

import javax.persistence.*;
import java.util.Date;

@Data
@Entity(name = "TEMP_STAFF_INSURANCE")
public class TempStaffInsurance {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private int seq;
    private String empNo;
    private String empName;
    @Column(name = "INSURANCE_YYMM")
    private String insuranceYyMm;
    private Double hourRate;
    private Double hours;
    private Integer salary;
    private Integer laborInsurance;
    private Integer healthInsurance;
    private Integer retireInsurance;
    private Date updateDate;
    private String updateUser;

}
