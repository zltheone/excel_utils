package com.zl.excelutils.enums;

public enum ExcelValueConvertEnum {
    TRUE("是", "T"),
    FALSE("否", "F"),
    CONFIRM("确认", "CONFIRM"),
    CANCEL("取消", "CANCEL"),
    ;

    private String key;

    private String value;

    ExcelValueConvertEnum(String key, String value) {
        this.key = key;
        this.value = value;
    }

    public String getKey() {
        return key;
    }

    public String getValue() {
        return value;
    }
}
