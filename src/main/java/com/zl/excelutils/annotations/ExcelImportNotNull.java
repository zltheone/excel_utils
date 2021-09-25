package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelImportNotNull {
    /**
     * 用于Excel导入，校验此属性不能为空
     **/
    boolean value() default true;
}
