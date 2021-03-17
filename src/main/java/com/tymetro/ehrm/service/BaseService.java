package com.tymetro.ehrm.service;

import java.util.List;
import java.util.Optional;

public interface BaseService<R, Long> {

    List<R> findAll();

    Optional<R> findById(Long l);

    R saveOrUpdate(R r);

    List<R> saveALL(List<R> list);

    void delete(R r);

    void deleteById(Long l);

    long count();

    boolean existsById(Long l);

}
