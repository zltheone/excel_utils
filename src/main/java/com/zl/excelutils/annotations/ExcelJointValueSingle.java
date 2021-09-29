package com.zl.excelutils.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelJointValueSingle {
    /**
     * 用于导入
     * value为单个值时,判断当前属性是否唯一
     * value为"属性名1,属性名2,..."拼接的多个值时，判断多个属性组合的值是否唯一
     * 备注：如果有重复，value内容会作为报错的表头提示出来
     **/
    String value() default "";

}
