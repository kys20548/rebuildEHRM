package com.tymetro.ehrm.utils;

import java.math.BigDecimal;

public class CHECKUTIL {


    public static BigDecimal nonNullBigDecimal(Object obj) {
        if (obj == null) {
            return new BigDecimal("0");
        }
        return (BigDecimal) obj;
    }
}
