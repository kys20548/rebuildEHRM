package com.tymetro.ehrm.controller;

import com.tymetro.ehrm.model.TempStaffInsurance;
import com.tymetro.ehrm.model.TempStaffSalary;
import com.tymetro.ehrm.repository.TempStaffInsuranceRepository;
import com.tymetro.ehrm.service.TempStaffInsuranceService;
import com.tymetro.ehrm.service.TempStaffSalaryService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
public class testController {

    @Autowired
    TempStaffInsuranceService tempStaffInsuranceService;

    @Autowired
    TempStaffSalaryService tempStaffSalaryService;

    @RequestMapping(path = "/test")
    public String test() {
        return tempStaffInsuranceService.count() + "";
    }

    @RequestMapping(path = "/test1")
    public String test1() {
        List<TempStaffInsurance> insu = tempStaffInsuranceService.findAll();
        insu.forEach(System.out::println);

        return insu.size()+"";
    }

    @RequestMapping(path = "/test2")
    public String test2() {
        List<TempStaffSalary> salarys = tempStaffSalaryService.findAll();
        salarys.forEach(System.out::println);

        return salarys.size()+"";
    }
}
