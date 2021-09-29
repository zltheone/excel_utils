package com.zl.excelutils.annotations;


import com.zl.excelutils.enums.ExcelValueConvertEnum;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelValueConvert {
    /**
     * 用于导入，该注解指定枚举将单元格内容转换成其他指定内容
     **/
    ExcelValueConvertEnum[] value();
}
