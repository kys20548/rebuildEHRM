package com.tymetro.ehrm.model;

import lombok.*;

import javax.persistence.*;
import java.util.Date;

@Getter
@Setter
@EqualsAndHashCode
@Entity
public class TempStaffInsurance {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private int id;
    private String empNo;
    private String empName;
    private String insuranceYYMM;
    private Double hourRate;
    private Double hours;
    private Integer salary;
    private Integer laborInsu;
    private Integer healthInsu;
    private Integer retireInsu;
    private Date updateDate;
    private String updateUser;

}
