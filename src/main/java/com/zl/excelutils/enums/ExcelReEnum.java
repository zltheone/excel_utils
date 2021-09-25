package com.zl.excelutils.enums;
public enum ExcelReEnum {
    YEAR2_WEEK2("^\\d{4}$", "年份后两位+周次 例：2123"),
    YEAR2_WEEK2_A_OR_B("^\\d{4}[A|B]$", "年份后两位+周次+A/B  例：2123A"),
    PHONE_NUMBER("^1\\d{10}", "手机号"),
    EMAIL("^\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*$", "邮箱"),
    ;

    private String re;

    private String desc;

    ExcelReEnum(String re, String example) {
        this.re = re;
        this.desc = example;
    }

    public String getDesc() {
        return desc;
    }

    public String getRe() {
        return re;
    }
}
