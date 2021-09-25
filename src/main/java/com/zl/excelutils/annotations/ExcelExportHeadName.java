package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelExportHeadName {
    /**
     * 用于导出，加在属性上，对应表头的名称
     * 属性上不加此注解默认不导出(忽略属性)
     * 调整导出的Excel顺序，只需调整对象的顺序即可
     **/
    String value() default "";
}
