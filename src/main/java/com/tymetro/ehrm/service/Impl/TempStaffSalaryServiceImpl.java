package com.tymetro.ehrm.service.Impl;

import com.tymetro.ehrm.model.TempStaffSalary;
import com.tymetro.ehrm.repository.TempStaffSalaryRepository;
import com.tymetro.ehrm.service.TempStaffSalaryService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;

@Service
public class TempStaffSalaryServiceImpl implements TempStaffSalaryService<TempStaffSalary, Long> {

    @Autowired
    TempStaffSalaryRepository tempStaffSalaryRepository;

    @Override
    public List<TempStaffSalary> findAll() {
        return tempStaffSalaryRepository.findAll();
    }

    @Override
    public Optional<TempStaffSalary> findById(Long l) {
        return tempStaffSalaryRepository.findById(l);
    }

    @Override
    public TempStaffSalary saveOrUpdate(TempStaffSalary tempStaffSalary) {
        return tempStaffSalaryRepository.save(tempStaffSalary);
    }

    @Override
    public List<TempStaffSalary> saveALL(List<TempStaffSalary> list) {
        return tempStaffSalaryRepository.saveAll(list);
    }

    @Override
    public void delete(TempStaffSalary tempStaffSalary) {
        tempStaffSalaryRepository.delete(tempStaffSalary);
    }

    @Override
    public void deleteById(Long l) {
        tempStaffSalaryRepository.deleteById(l);
    }

    @Override
    public long count() {
        return tempStaffSalaryRepository.count();
    }

    @Override
    public boolean existsById(Long l) {
        return tempStaffSalaryRepository.existsById(l);
    }
}
