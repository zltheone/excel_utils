package com.zl.excelutils.domain;

import com.zl.excelutils.annotations.ExcelExportHeadName;
import com.zl.excelutils.annotations.ExcelSheetName;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.util.CollectionUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Description: excel导出工具入参对象
 * @Date: 2021/8/26 16:55
 */
@Getter
@Setter
@NoArgsConstructor
public class ExcelVO<T> {

    /** 表头名称列表 */
    private List<String> headNameList;

    /** 数据的对象集合 */
    private List<T> data;

    /** 对象属性列表 */
    private List<Field> fieldsList;

    /** 创建的Sheet名字 */
    private String sheetName;


    public ExcelVO(List<T> data){
        if (CollectionUtils.isEmpty(data)){
            return;
        }
        T t = data.get(0);
        this.headNameList = getExcelHeadNameList(t);
        this.fieldsList = getActualFields(t);
        this.sheetName = getExcelSheetName(t);
        this.data = data;
    }


    /**
     * 获取类的所有属性列表
     * 只获取加了@ExcelHeadName注解的属性
     **/
    private List<Field> getActualFields(T t) {
        List<Field> fieldsList = new ArrayList<>(Arrays.asList(t.getClass().getDeclaredFields()));
        return fieldsList.stream().filter(f -> f.getAnnotation(ExcelExportHeadName.class) != null).collect(Collectors.toList());
    }

    /**
     * 获取Excel表头名称列表
     **/
    private List<String> getExcelHeadNameList(T t) {
        ArrayList<String> headNameList = new ArrayList<>();
        getActualFields(t).forEach(s -> headNameList.add(s.getAnnotation(ExcelExportHeadName.class).value()));
        return headNameList;
    }

    /**
     * 获取Excel SheetName
     **/
    private String getExcelSheetName(T t) {
        if (t.getClass().isAnnotationPresent(ExcelSheetName.class)) {
            return t.getClass().getAnnotation(ExcelSheetName.class).value();
        }
        return "sheet";
    }
}
