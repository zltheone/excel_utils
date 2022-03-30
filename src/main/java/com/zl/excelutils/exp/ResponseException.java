package com.zl.excelutils.exp;


/**
 * @Description: 自定义业务异常类
 * @Date: 2021/11/24 19:20
 */
public class ResponseException extends RuntimeException {

    private Integer code;

    private String msg;

    public ResponseException(){
        super();
    }

    public ResponseException(Integer code, String msg){
        super(msg);
        this.code = code;
        this.msg = msg;
    }

    public ResponseException(String msg){
        super(msg);
    }

    public Integer getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }
}
