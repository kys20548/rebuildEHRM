package com.tymetro.ehrm.controller;

import com.tymetro.ehrm.repository.TempStaffInsuranceRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class testController {
    @Autowired
    TempStaffInsuranceRepository tempStaffInsuranceRepository;
    @RequestMapping(path="/test")
    @ResponseBody
    public String xxx(){
        return tempStaffInsuranceRepository.count()+"";
    }
}
