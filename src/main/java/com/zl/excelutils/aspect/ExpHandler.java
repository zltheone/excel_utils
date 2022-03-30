package com.zl.excelutils.aspect;

import com.zl.excelutils.exp.ExcelUtilsException;
import com.zl.excelutils.exp.ResponseException;
import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.annotation.AfterThrowing;
import org.aspectj.lang.annotation.Aspect;

/**
 * @Description: 异常处理类
 * @Date: 2021/11/23 16:24
 */
@Aspect
@Slf4j
public class ExpHandler {

    /**
     * 打印Excel工具类异常
     **/
    @AfterThrowing(pointcut = "execution(* com.zl.excelutils..*.*(..))", throwing = "e")
    public void ExcelUtilsExceptionHandler(ExcelUtilsException e) {
        log.error("ExcelUtils工具类异常", e);
    }


    /**
     * 打印响应工具类异常
     **/
    @AfterThrowing(pointcut = "execution(* com.zl.excelutils..*.*(..))", throwing = "e")
    public void ResponseExceptionHandler(ResponseException e) {
        log.error("ResponseUtils响应工具类异常", e);
    }
}
