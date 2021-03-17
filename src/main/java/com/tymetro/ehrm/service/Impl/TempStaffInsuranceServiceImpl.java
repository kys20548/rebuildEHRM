package com.tymetro.ehrm.service.Impl;

import com.tymetro.ehrm.model.TempStaffInsurance;
import com.tymetro.ehrm.repository.TempStaffInsuranceRepository;
import com.tymetro.ehrm.service.TempStaffInsuranceService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;

@Service
public class TempStaffInsuranceServiceImpl implements TempStaffInsuranceService<TempStaffInsurance, Long> {

    @Autowired
    TempStaffInsuranceRepository tempStaffInsuranceRepository;

    @Override
    public List<TempStaffInsurance> findAll() {
        return tempStaffInsuranceRepository.findAll();
    }

    @Override
    public Optional<TempStaffInsurance> findById(Long l) {
        return tempStaffInsuranceRepository.findById(l);
    }

    @Override
    public TempStaffInsurance saveOrUpdate(TempStaffInsurance tempStaffInsurance) {
        return tempStaffInsuranceRepository.save(tempStaffInsurance);
    }

    @Override
    public List<TempStaffInsurance> saveALL(List<TempStaffInsurance> list) {
        return tempStaffInsuranceRepository.saveAll(list);
    }

    @Override
    public void delete(TempStaffInsurance tempStaffInsurance) {
        tempStaffInsuranceRepository.delete(tempStaffInsurance);
    }

    @Override
    public void deleteById(Long l) {
        tempStaffInsuranceRepository.deleteById(l);
    }

    @Override
    public long count() {
        return tempStaffInsuranceRepository.count();
    }

    @Override
    public boolean existsById(Long l) {
        return tempStaffInsuranceRepository.existsById(l);
    }
}
