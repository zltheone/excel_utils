package com.zl.excelutils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.zl.excelutils.domain.AdsOcLackMatFillInfoPO;
import com.zl.excelutils.utils.ExcelUtils;
import org.apache.commons.collections.CollectionUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @Date: 2022/2/24 18:27
 */
@RunWith(SpringRunner.class)
@SpringBootTest(classes = ExcelUtilsTest.class)
public class ExcelUtilsTest {


    @Test
    public void testImport(){
        // String filePath = "C:\\Users\\li5.zhong\\Desktop\\TCL\\需求\\OC\\齐套\\测试导入表\\欠料补料 - mistake.xlsx"; // 校验会报错的表
        String filePath = "C:\\Users\\li5.zhong\\Desktop\\TCL\\需求\\OC\\齐套\\测试导入表\\欠料补料 - right.xlsx"; // 能正确导入的表
        // 获取Excel数据信息
        Map<String, Object> map = ExcelUtils.importExcelToList(filePath, AdsOcLackMatFillInfoPO.class);
        if (map.get("error") != null) {
            List<String> error = JSONArray.parseArray(JSONObject.toJSONString(map.get("error"))).toJavaList(String.class);
            if (CollectionUtils.isNotEmpty(error)) {  // 有报错直接返回提示信息
                error.forEach(System.out::println);
            }
        }
        if (map.get("data") != null) {
            List<AdsOcLackMatFillInfoPO> data = JSONArray.parseArray(JSONObject.toJSONString(map.get("data")), AdsOcLackMatFillInfoPO.class);
            data.forEach(System.out::println);
            if (CollectionUtils.isNotEmpty(data)) {
                // 打开批处理,提交到数据库
            }
        }
    }

    @Test
    public void testExportMap(){

    }


}
