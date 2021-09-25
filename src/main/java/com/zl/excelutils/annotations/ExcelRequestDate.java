package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelRequestDate {
    /**
     * 用于导入，校验此属性必须为某种日期格式
     * 可传不同值，如：yyyy-MM， yyyy-MM-dd HH:mm， yyyy-MM-dd HH:mm:ss等
     **/
    String value() default "yyyy-MM-dd";
}
