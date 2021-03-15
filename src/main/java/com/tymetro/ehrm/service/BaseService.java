package com.tymetro.ehrm.service;

import java.util.List;
import java.util.Optional;

public interface BaseService<R,Long> {

    public List<R> findAll();

    public Optional<R> findById(Long l);

    public R saveOrUpdate(R r);

    public List<R> saveALL(List<R> list);

    public void delete(R r);

    public void deleteById(Long l);

    public long count();

    public boolean existsById(Long l);

}
