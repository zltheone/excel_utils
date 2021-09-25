package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelImportIgnoreField {
    /**
     * 用于导入加在忽略的属性上(即：Excel表中没有的属性)
     **/
    String value() default "";
}
