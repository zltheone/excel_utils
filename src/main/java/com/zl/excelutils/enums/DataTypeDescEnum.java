package com.zl.excelutils.enums;

public enum DataTypeDescEnum {

    BIG_DECIMAL_DESC("java.math.BigDecimal", "数字类型"),
    INTEGER("java.lang.Integer", "数字类型"),
    LONG("java.lang.Long", "数字类型"),
    DOUBLE("java.lang.Double", "数字类型"),
    SHORT("java.lang.Short", "数字类型"),
    STRING("java.lang.String", "文本类型"),
    BOOLEAN("java.lang.Boolean", "True Or False"),
    DATE("java.util.Date", "日期类型"),
    ;

    private String code;

    private String desc;

    DataTypeDescEnum(String code, String desc) {
        this.code = code;
        this.desc = desc;
    }

    public String getCode() {
        return code;
    }

    public String getDesc() {return desc; }

    public static String getDescByCode(String code) {
        for (DataTypeDescEnum type : DataTypeDescEnum.values()){
            if (type.getCode().equals(code)){
                return type.desc;
            }
        }
        return "其他类型";
    }
}
