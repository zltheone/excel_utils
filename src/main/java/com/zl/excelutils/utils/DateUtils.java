package com.zl.excelutils.utils;

import cn.hutool.core.date.DateUtil;
import com.zl.excelutils.exp.ExcelUtilsException;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Description:
 * @Date: 2021/11/5 16:56
 */
public class DateUtils {

    private static final String[] parsePatterns = {
            "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM", "yyyy-MM-dd HH:mm:ss.SSS",
            "yyyy/MM/dd", "yyyy/MM/dd HH:mm:ss", "yyyy/MM/dd HH:mm", "yyyy/MM",
            "yyyy.MM.dd", "yyyy.MM.dd HH:mm:ss", "yyyy.MM.dd HH:mm", "yyyy.MM",
            "yyyy年MM月dd日", "yyyy年MM月dd日 HH时mm分ss秒", "yyyy年MM月dd日 HH时mm分", "yyyy年MM月",
            "yyyyMMdd", "yyyyMMddHHmmss", "yyyyMMddHHmm", "yyyyMM"};


    /**
     * 将字符串格式化为指定格式
     **/
    public static String formatDate(String value, String pattern){
        if (StringUtils.isBlank(value) || StringUtils.isBlank(pattern)){
            return "";
        }
        Date date = DateUtil.parse(value, parsePatterns);
        return DateUtil.format(date, pattern);
    }

    /**
     * 将字符串格式化为指定日期格式
     **/
    public static Date formatStrToDate(String value, String pattern){
        try {
            SimpleDateFormat format = new SimpleDateFormat(pattern);
            return format.parse(value);
        }catch (ParseException e){
            throw new ExcelUtilsException(-1, "转换日期异常");
        }
    }


    /**
     * 将字符串格式化为指定日期格式
     **/
    public static String formatDateToStr(Date date, String pattern){
        SimpleDateFormat format = new SimpleDateFormat(pattern);
        return format.format(date);
    }

    /**
     * 设置对象属性值
     * param: o: 对象, fieldName: 属性名, value: 值
     **/
    public static void setFieldValueByName(Object o, String fieldName, Object value) throws InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        String firstLetter = fieldName.substring(0, 1).toUpperCase();
        String setter = "set" + firstLetter + fieldName.substring(1);
        Method method = o.getClass().getMethod(setter, value.getClass());
        method.invoke(o, value);
    }

    /**
     * 处理开始时间，拼接上 00:00:00
     **/
    public static String handleStartDate(String date){
        if (isDate(date, "yyyy-MM-dd")){
            return date + " 00:00:00";
        }
        throw new ExcelUtilsException(-1, "日期格式不正确");
    }

    /**
     * 处理结束时间，拼接上 23:59:59
     **/
    public static String handleEndDate(String date){
        if (isDate(date, "yyyy-MM-dd")){
            return date + " 23:59:59";
        }
        throw new ExcelUtilsException(-1, "日期格式不正确");
    }


    /**
     * 判断是否为指定日期格式
     **/
    public static boolean isDate(String date, String pattern){
        SimpleDateFormat format = new SimpleDateFormat(pattern);
        ParsePosition pos = new ParsePosition(0);
        format.setLenient(false);
        Date result = format.parse(date, pos);
        return !(pos.getIndex() == 0) && date.equals(format.format(result));
    }


}
