package com.zl.excelutils.exp;

/**
 * @Description: Excel工具类异常
 * @Date: 2021/11/23 16:19
 */
public class ExcelUtilsException extends RuntimeException {

    private String msg;

    private Integer code;

    public ExcelUtilsException(){
        super();
    }

    public ExcelUtilsException(String message){
        super(message);
    }

    public ExcelUtilsException(Integer code, String msg){
        super(msg);
        this.code = code;
        this.msg = msg;
    }

    public String getMsg() {
        return msg;
    }

    public Integer getCode() {
        return code;
    }
}
