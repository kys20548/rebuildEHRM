package com.tymetro.ehrm.vo;

import java.math.BigDecimal;

public class TotalSalaryVO {
    private int countNO;
    private BigDecimal salaryTotal;
    private BigDecimal laborSelfTotal;
    private BigDecimal laborComTotal;
    private BigDecimal retireSelfTotal;
    private BigDecimal retireComTotal;
    private BigDecimal healthSelfTotal;
    private BigDecimal healthComTotal;
    private BigDecimal repaymentFundTotal;
    private BigDecimal welfare;
    private BigDecimal salaryFinalTotal;

    public int getCountNO() {
        return countNO;
    }
    public void setCountNO(int countNO) {
        this.countNO = countNO;
    }
    public BigDecimal getSalaryTotal() {
        return salaryTotal;
    }
    public void setSalaryTotal(BigDecimal salaryTotal) {
        this.salaryTotal = salaryTotal;
    }
    public BigDecimal getLaborSelfTotal() {
        return laborSelfTotal;
    }
    public void setLaborSelfTotal(BigDecimal laborSelfTotal) {
        this.laborSelfTotal = laborSelfTotal;
    }
    public BigDecimal getLaborComTotal() {
        return laborComTotal;
    }
    public void setLaborComTotal(BigDecimal laborComTotal) {
        this.laborComTotal = laborComTotal;
    }
    public BigDecimal getRetireSelfTotal() {
        return retireSelfTotal;
    }
    public void setRetireSelfTotal(BigDecimal retireSelfTotal) {
        this.retireSelfTotal = retireSelfTotal;
    }
    public BigDecimal getRetireComTotal() {
        return retireComTotal;
    }
    public void setRetireComTotal(BigDecimal retireComTotal) {
        this.retireComTotal = retireComTotal;
    }
    public BigDecimal getHealthSelfTotal() {
        return healthSelfTotal;
    }
    public void setHealthSelfTotal(BigDecimal healthSelfTotal) {
        this.healthSelfTotal = healthSelfTotal;
    }
    public BigDecimal getHealthComTotal() {
        return healthComTotal;
    }
    public void setHealthComTotal(BigDecimal healthComTotal) {
        this.healthComTotal = healthComTotal;
    }
    public BigDecimal getRepaymentFundTotal() {
        return repaymentFundTotal;
    }
    public void setRepaymentFundTotal(BigDecimal repaymentFundTotal) {
        this.repaymentFundTotal = repaymentFundTotal;
    }
    public BigDecimal getSalaryFinalTotal() {
        return salaryFinalTotal;
    }
    public void setSalaryFinalTotal(BigDecimal salaryFinalTotal) {
        this.salaryFinalTotal = salaryFinalTotal;
    }
    public BigDecimal getWelfare() {
        return welfare;
    }
    public void setWelfare(BigDecimal welfare) {
        this.welfare = welfare;
    }

    public void initValue() {
        countNO=0;
        salaryTotal=BigDecimal.ZERO;
        laborSelfTotal=BigDecimal.ZERO;
        laborComTotal=BigDecimal.ZERO;
        retireSelfTotal=BigDecimal.ZERO;
        retireComTotal=BigDecimal.ZERO;
        healthSelfTotal=BigDecimal.ZERO;
        healthComTotal=BigDecimal.ZERO;
        repaymentFundTotal=BigDecimal.ZERO;
        salaryFinalTotal=BigDecimal.ZERO;
        welfare=BigDecimal.ZERO;

    }

}
