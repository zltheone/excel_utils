package com.zl.excelutils.annotations;


import com.zl.excelutils.enums.ExcelReEnum;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelReValue {
    /**
     * 用于导入时，校验一些特殊格式的属性，可自定义正则格式
     * 传入Excel格式校验枚举数组
     **/
    ExcelReEnum[] value();
}
