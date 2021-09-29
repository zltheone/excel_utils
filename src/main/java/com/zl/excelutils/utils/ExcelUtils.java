package com.zl.excelutils.utils;

import cn.hutool.core.convert.Convert;
import cn.hutool.core.date.DateException;
import cn.hutool.core.date.DateTime;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.ReUtil;
import com.zl.excelutils.annotations.*;
import com.zl.excelutils.domain.CellValueVO;
import com.zl.excelutils.domain.ExcelVO;
import com.zl.excelutils.enums.DataTypeDescEnum;
import com.zl.excelutils.enums.ExcelReEnum;
import com.zl.excelutils.enums.ExcelValueConvertEnum;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Description: POI Excel工具类，导入导出都为List<PO>
 * @Date: 2021/9/17 11:39
 */
@Slf4j
public class ExcelUtils {

    private static final Integer ROW_ACCESS_WINDOW_SIZE = 100;

    private static final String[] parsePatterns = {
            "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM",
            "yyyy/MM/dd", "yyyy/MM/dd HH:mm:ss", "yyyy/MM/dd HH:mm", "yyyy/MM",
            "yyyy.MM.dd", "yyyy.MM.dd HH:mm:ss", "yyyy.MM.dd HH:mm", "yyyy.MM",
            "yyyy年MM月dd日", "yyyy年MM月dd日 HH时mm分ss秒", "yyyy年MM月dd日 HH时mm分", "yyyy年MM月",
            "yyyyMMdd", "yyyyMMddHHmmss", "yyyyMMddHHmm", "yyyyMM"};

    /**
     * 导出Excel工具(一个Sheet使用)
     * param: excelVO : excel导出工具入参对象, 一个Sheet对应一个ExcelVO
     * return 字节数组
     **/
    public static <T> byte[] exportExcel(ExcelVO<T> excelVO) {
        byte[] bytes = null;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            // 创建工作簿
            SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
            createSheetData(workbook.getXSSFWorkbook(), excelVO);
            // 创建工作表
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        } catch (Exception e) {
            log.error("导出Excel异常 {}", e);
        } finally {
            try {
                outputStream.flush();
                outputStream.close();
            } catch (Exception e) {
                log.error("exportExcel关闭流异常 {}", e);
            }
        }
        return bytes;
    }


    /**
     * 导出Excel工具(多个sheet使用)
     * param: excelVOS : excel工具对象列表, 一个Sheet对应一个ExcelVO
     * return 字节数组
     **/
    public static byte[] exportExcel(List<ExcelVO> excelVOS) {
        byte[] bytes = null;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            // 创建工作簿
            SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), ROW_ACCESS_WINDOW_SIZE);
            for (ExcelVO excelVO : excelVOS) {
                createSheetData(workbook.getXSSFWorkbook(), excelVO);
            }
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        } catch (Exception e) {
            log.error("导出Excel异常 {}", e);
        } finally {
            try {
                outputStream.flush();
                outputStream.close();
            } catch (Exception e) {
                log.error("exportExcel关闭流异常 {}", e);
            }
        }
        return bytes;
    }


    /**
     * 输出工作簿到指定地址
     * param: excelVO : excel工具对象, 一个Sheet对应一个ExcelVO
     * param: filePath : 本地下载excel文件路径地址， 可为null
     **/
    public static void downloadExcel(ExcelVO excelVO, String filePath) {
        try {
            FileOutputStream outputStream = new FileOutputStream(filePath);
            // 创建工作簿
            SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), ROW_ACCESS_WINDOW_SIZE);
            createSheetData(workbook.getXSSFWorkbook(), excelVO);
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            log.error("导出Excel到{}异常， {}", filePath, e);
        }
    }


    /**
     * 输出工作簿到指定地址
     * param: excelVOS : excel工具对象列表, 一个Sheet对应一个ExcelVO
     * param: filePath : 本地下载excel文件路径地址， 可为null
     **/
    public static void downloadExcel(List<ExcelVO> excelVOS, String filePath) {
        try {
            FileOutputStream outputStream = new FileOutputStream(filePath);
            // 创建工作簿
            SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), ROW_ACCESS_WINDOW_SIZE);
            for (ExcelVO excelVO : excelVOS) {
                createSheetData(workbook.getXSSFWorkbook(), excelVO);
            }
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            log.error("导出Excel到{}异常， {}", filePath, e);
        }
    }


    /**
     * 导入Excel，保存为List对象列表
     * 配合自定义注解校验导入格式，导入的Excel表头需与属性顺序保持一致
     * Param filePath 文件路径
     * Param T: 应用对象
     * Return Map<String, Object> {data: 导入的数据, error: 错误提示, header: 表头名称}
     **/
    public static <T> Map<String, Object> importExcelToList(String filePath, Class<T> tClass) {
        try {
            InputStream stream = new FileInputStream(filePath);
            SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(stream), 100);
            XSSFSheet sheet = workbook.getXSSFWorkbook().getSheetAt(0);
            return getSheetData(tClass, sheet);
        } catch (Exception e) {
            log.error("从{}导入Excel异常， {}", filePath, e);
        }
        return new HashMap<>();
    }


    /**
     * 导入Excel，保存为List对象列表
     * 配合自定义注解校验导入格式，导入的Excel表头需与属性顺序保持一致
     * Param bytes 字节数组
     * Param T: 应用对象
     * Return Map<String, Object> {data: 导入的数据, error: 错误提示, header: 表头名称}
     **/
    public static <T> Map<String, Object> importExcelToList(byte[] bytes, Class<T> tClass) {
        try {
            InputStream stream = new ByteArrayInputStream(bytes);
            SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(stream), 100);
            XSSFSheet sheet = workbook.getXSSFWorkbook().getSheetAt(0);
            return getSheetData(tClass, sheet);
        } catch (Exception e) {
            log.error("导入Excel异常， {}", e);
        }
        return new HashMap<>();
    }


    /**
     * 获取Sheet数据
     **/
    private static <T> Map<String, Object> getSheetData(Class<T> tClass, XSSFSheet sheet) throws IllegalAccessException, InstantiationException {
        Map<String, Object> result = new HashMap<>();
        List<T> data = new ArrayList<>();
        List<String> errorList = new ArrayList<>();
        // 获取对象的属性名列表
        List<Field> fieldsList = getExcelImportFields(tClass);
        StringBuilder header = new StringBuilder();
        Map<String, List<Object>> singleValueMap = new HashMap();
        // 获取最后一行的num，即总行数
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            // 判断是否为空行
            if (isNoneRow(row)) {
                continue;
            }
            if (i == 0) {
                row.forEach(r -> {    // 获取表头
                    if (header.length() > 0) {
                        header.append(",");
                    }
                    header.append(r.getStringCellValue().trim());
                });
                // 校验表头
                String fieldHeaders = getFieldHeaders(fieldsList);
                if (!fieldHeaders.equals(header.toString())) {
                    errorList.add("表头与模板不一致，请重新下载模板");
                    break;
                }
                continue;
            }
            // 使用反射创建对象
            Object o = tClass.newInstance();
            // 行map, 用于拼接多个联合属性判断的值
            HashMap<String, String> map = new HashMap<>();
            for (int rol = 0; rol < fieldsList.size(); rol++) {
                // 获取当列的值
                Object value = getCellValue(row.getCell(rol));
                Field field = fieldsList.get(rol);
                // 校验单元格内容并获取格式化后的返回值和错误提示
                CellValueVO cellValueVO = verifyField(field, value);
                if (StringUtils.isNotBlank(cellValueVO.getCellError())) {
                    errorList.add("第" + (i + 1) + "行" + "第" + (rol + 1) + "列【" + sheet.getRow(0).getCell(rol).getStringCellValue() + "】" + cellValueVO.getCellError());
                    continue;
                }
                // 调用对象的set方法构建对象
                if (cellValueVO.getCellValue() != null) {
                    setFieldValueByName(o, field.getName(), cellValueVO.getCellValue());
                }
                // 拼接联合判断值的属性内容，空不校验
                if (field.getAnnotation(ExcelJointValueSingle.class) != null && cellValueVO.getCellValue() != null && StringUtils.isNotBlank(cellValueVO.getCellValue().toString())) {
                    String key = field.getAnnotation(ExcelJointValueSingle.class).value();
                    if (map.containsKey(key)) {
                        String s = map.get(key);
                        s += cellValueVO.getCellValue().toString();
                        map.put(key, s);
                        continue;
                    }
                    map.put(key, cellValueVO.getCellValue().toString());
                }
            }
            if (!map.isEmpty()) {
                // 校验联合判断的属性内容是否重复
                for (String key : map.keySet()) {
                    // 是否已存在联合属性的key
                    if (!singleValueMap.containsKey(key)) {
                        List<Object> objects = new ArrayList<>();
                        objects.add(map.get(key));
                        singleValueMap.put(key, objects);
                        continue;
                    }
                    // 是否包含联合属性内容
                    if (singleValueMap.get(key).contains(map.get(key))) {
                        errorList.add("第" + (i + 1) + "行【" + key + "】存在重复内容");
                        continue;
                    }
                    singleValueMap.get(key).add(map.get(key));
                }
            }
            data.add((T) o);
        }
        result.put("header", header.toString());
        result.put("data", data);
        result.put("error", errorList);
        return result;
    }


    /**
     * 判断单元格内容是否正确
     * Param field: 属性  cellValue: 单元格值
     * 使用注解实现对每个属性的不同校验
     **/
    private static CellValueVO verifyField(Field field, Object cellValue) {
        CellValueVO vo = new CellValueVO();
        ArrayList<String> errorList = new ArrayList<>();
        Annotation[] annotations = field.getAnnotations();
        String cs = (cellValue != null ? cellValue.toString() : "").trim();
        for (Annotation annotation : annotations) {
            String type = annotation.annotationType().getSimpleName();
            if (type.equals("ExcelExportHeadName")) {
                continue;
            }
            if (type.equals("ExcelImportNotNull") && StringUtils.isBlank(cs)) {
                errorList.add("不能为空");
                continue;
            }
            // 只判断属性指定条件，单元格可为空，空不返回错误
            if (type.equals("ExcelAssignValue") && StringUtils.isNotBlank(cs)) {
                String value = field.getAnnotation(ExcelAssignValue.class).value();
                String[] strings = value.split(",");
                if (!Arrays.asList(strings).contains(cs)) {
                    errorList.add("必须在指定范围【" + value + "】内");
                }
                continue;
            }
            // 只判断日期格式，单元格可为空，空不返回错误
            if (type.equals("ExcelRequestDate") && StringUtils.isNotBlank(cs)) {
                String dateFormat = field.getAnnotation(ExcelRequestDate.class).value();
                try {
                    // 判断是否为日期后转换成指定的格式
                    DateTime date = DateUtil.parse(cs, parsePatterns);
                    cellValue = DateUtil.format(date, dateFormat);
                } catch (DateException e) {
                    errorList.add("日期格式必须为【" + dateFormat + "】");
                }
                continue;
            }
            // 只判断正则，可为正则数组，满足其一不报错，单元格可为空，空不返回错误
            if (type.equals("ExcelReValue") && StringUtils.isNotBlank(cs)) {
                ExcelReEnum[] reEnums = field.getAnnotation(ExcelReValue.class).value();
                if (Arrays.stream(reEnums).noneMatch(r -> ReUtil.isMatch(r.getRe(), cs))) {
                    for (int i = 0; i < reEnums.length; i++) {
                        StringBuilder builder = new StringBuilder();
                        if (i == 0) {
                            builder.append("格式不正确,必须为【").append(reEnums[i].getDesc()).append("】");
                        } else {
                            builder.append("或者【").append(reEnums[i].getDesc()).append("】");
                        }
                        errorList.add(builder.toString());
                    }
                }
                continue;
            }
            // 判断内容是否需要转换, 单元格可为空，空不返回错误
            if (type.equals("ExcelValueConvert") && StringUtils.isNotBlank(cs)) {
                ExcelValueConvertEnum[] values = field.getAnnotation(ExcelValueConvert.class).value();
                boolean flag = true;
                for (ExcelValueConvertEnum convertEnum : values) {
                    if (cs.equals(convertEnum.getKey())) {
                        cellValue = convertEnum.getValue();
                        flag = false;
                    }
                }
                if (flag) {
                    errorList.add("【" + cs + "】" + "找不到对应格式的转换");
                }
            }
        }
        // 将通过验证后的单元格数据类型转换属性所需的数据类型
        String typeName = field.getGenericType().getTypeName();
        if (cellValue != null && !typeName.equals(cellValue.getClass().getName())) {
            try {
                cellValue = Convert.convert(field.getGenericType(), cellValue);
            } catch (Exception e) {
                errorList.add("单元格格式不正确,请输入【" + DataTypeDescEnum.getDescByCode(typeName) + "】");
                log.error("转换属性【{}】类型【{}】To 【{}】,单元格内容【{}】, 异常:{}", field.getName(), cellValue.getClass().getName(), typeName, cellValue, e);
            }
        }
        vo.setCellError(StringUtils.join(errorList, ","));
        vo.setCellValue(cellValue);
        return vo;
    }


    /**
     * 判断是否为空行
     **/
    private static boolean isNoneRow(XSSFRow row) {
        if (row == null) {
            return true;
        }
        // 判断空行 这里加了格式(如边框线等)的行不会被判断为null，需要另加判断所有列是否为空
        for (int rol = 0; rol <= row.getLastCellNum(); rol++) {
            XSSFCell cell = row.getCell(rol);
            Object value = getCellValue(cell);
            if (cell != null && value != null && StringUtils.isNotBlank(value.toString())) {
                return false;
            }
        }
        return true;
    }


    /**
     * 设置对象属性值
     * param: o: 对象, fieldName: 属性名, value: 值
     **/
    private static void setFieldValueByName(Object o, String fieldName, Object value) {
        try {
            String firstLetter = fieldName.substring(0, 1).toUpperCase();
            String setter = "set" + firstLetter + fieldName.substring(1);
            Method method = o.getClass().getMethod(setter, value.getClass());
            method.invoke(o, value);
        } catch (Exception e) {
            log.error("给对象【{}】的属性【{}】,设置【{}】,失败{}", o, fieldName, value, e);
        }
    }


    /**
     * 获取对象属性值
     * param: fieldName: 属性名， o: 对象
     **/
    private static Object getFieldValueByName(Object o, String fieldName) {
        try {
            String firstLetter = fieldName.substring(0, 1).toUpperCase();
            String getter = "get" + firstLetter + fieldName.substring(1);
            Method method = o.getClass().getMethod(getter, new Class[]{});
            return method.invoke(o);
        } catch (Exception e) {
            log.error("获取对象【{}】的属性【{}】值,失败{}", o, fieldName, e);
        }
        return null;
    }

    /**
     * 获取类需要设值的属性列表
     * param:  class 反射对象
     * 使用@ExcelImportIgnoreField 注解 忽略不需要设值的属性
     **/
    private static <T> List<Field> getExcelImportFields(Class<T> tClass) {
        List<Field> fieldsList = new ArrayList<>(Arrays.asList(tClass.getDeclaredFields()));
        // 去除忽略导入的属性
        List<Field> ignoreFieldsList = fieldsList.stream().filter(f -> f.getAnnotation(ExcelImportIgnoreField.class) != null).collect(Collectors.toList());
        fieldsList.removeAll(ignoreFieldsList);
        return fieldsList;
    }

    /**
     * 获取属性表头
     * param:  class 反射对象
     * 使用@ExcelImportIgnoreField 注解 忽略不需要设值的属性
     **/
    private static String getFieldHeaders(List<Field> fields) {
        StringBuilder builder = new StringBuilder();
        for (Field field : fields) {
            if (field.getAnnotation(ExcelExportHeadName.class) != null) {
                if (builder.length() > 0) {
                    builder.append(",");
                }
                builder.append(field.getAnnotation(ExcelExportHeadName.class).value());
            }
        }
        return builder.toString();
    }


    /**
     * 获取单元格内容
     * 判断单数据类型并返回对应数据
     **/
    private static Object getCellValue(XSSFCell cell) {
        Object cellValue = null;
        if (null != cell) {
            // 以下是判断数据的类型
            switch (cell.getCellTypeEnum()) {
                case NUMERIC: // 数字
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        cellValue = DateUtil.formatDate(date);
                    } else {
                        cellValue = new BigDecimal(cell.getNumericCellValue());
                    }
                    break;
                case STRING: // 字符串
                    cellValue = cell.getStringCellValue();
                    break;
                case BOOLEAN: // Boolean
                    cellValue = cell.getBooleanCellValue() + "";
                    break;
                case FORMULA: // 公式
                    cellValue = cell.getCellFormula() + "";
                    break;
                case BLANK: // 空值
                    cellValue = "";
                    break;
                case ERROR: // 故障
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
        }
        return cellValue;
    }


    /**
     * 创建Sheet 并写入数据
     **/
    private static <T> void createSheetData(XSSFWorkbook workbook, ExcelVO excelVO) {
        if (excelVOIsNotEmpty(excelVO)) {
            // 创建工作表
            XSSFSheet sheet = workbook.createSheet(excelVO.getSheetName());
            sheet.setDefaultColumnWidth((short) 15);
            // 居中且加边框的样式
            XSSFCellStyle borderCenterStyle = getBorderCenterStyle(workbook);
            // 只加边框的样式
            XSSFCellStyle borderStyle = getBorderStyle(workbook);
            // 表头样式
            XSSFCellStyle headerStyle = getHeaderStyle(workbook);

            List<T> data = excelVO.getData();
            // 创建第一行表头
            XSSFRow header = sheet.createRow(0);
            List<String> headerList = excelVO.getHeadNameList();
            for (int i = 0; i < headerList.size(); i++) {
                XSSFCell cell = header.createCell(i);
                cell.setCellStyle(headerStyle);
                cell.setCellValue(headerList.get(i));
                header.setHeightInPoints((short) 21); // 设置行高，单位：磅
            }
            // 获取对象的属性名列表
            List<Field> fieldsList = excelVO.getFieldsList();
            // 把list的数据写入到excel中
            for (int i = 0; i < data.size(); i++) {
                // 创建行，从第二行开始写入
                XSSFRow row = sheet.createRow(i + 1);
                //创建每个单元格Cell，即列的数据
                T po = data.get(i);
                for (int j = 0; j < fieldsList.size(); j++) {
                    XSSFCell cell = row.createCell(j);
                    Field field = fieldsList.get(j);
                    Object value = getFieldValueByName(po, field.getName());
                    String genericType = field.getGenericType().toString();
                    // 判断对象属性是否为数字类型, 设置单元格类型为数字格式
                    if (genericType.equals("class java.lang.Integer") || genericType.equals("class java.lang.Double") || genericType.equals("class java.math.BigDecimal")) {
                        cell.setCellStyle(borderCenterStyle);
                        cell.setCellValue(value == null ? 0 : Double.valueOf(String.valueOf(value)));
                    } else {
                        cell.setCellStyle(borderStyle);
                        cell.setCellValue(value == null ? "" : String.valueOf(value));
                    }
                }
            }
        }
    }

    /**
     * 判断vo是否为空
     **/
    private static boolean excelVOIsNotEmpty(ExcelVO excelVO) {
        return excelVO != null && CollectionUtils.isNotEmpty(excelVO.getData()) && CollectionUtils.isNotEmpty(excelVO.getFieldsList()) && CollectionUtils.isNotEmpty(excelVO.getHeadNameList()) && StringUtils.isNotBlank(excelVO.getSheetName());
    }

    /**
     * 设置表头样式
     **/
    private static XSSFCellStyle getHeaderStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        // font.setBold(true); // 加粗
        font.setFontName("微软雅黑"); // 字体
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充样式
        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 176, 80))); // 前景色
        setCellAllBorder(style);  // 加四周边框
        return style;
    }

    /**
     * 加四周边框且居中的样式
     **/
    private static XSSFCellStyle getBorderCenterStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        setCellAllBorder(style);  // 加四周边框
        style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        return style;
    }

    /**
     * 只加四周边框的样式
     **/
    private static XSSFCellStyle getBorderStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        setCellAllBorder(style);  // 加四周边框
        return style;
    }

    /**
     * 加四周边框
     **/
    private static void setCellAllBorder(XSSFCellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

}
