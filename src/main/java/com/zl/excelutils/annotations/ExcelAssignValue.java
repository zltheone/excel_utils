package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAssignValue {
    /**
     * 用于导入，校验自定义范围的内容
     * 如："1,-1",  "是,不是" 等等，使用","分隔范围内容
     **/
    String value() default "";
}
