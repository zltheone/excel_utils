package com.zl.excelutils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.zl.excelutils.domain.AdsOcLackMatFillInfoPO;
import com.zl.excelutils.domain.ExportVO;
import com.zl.excelutils.utils.DateUtils;
import com.zl.excelutils.utils.ExcelUtils;
import com.zl.excelutils.utils.ResponseUtils;
import org.apache.commons.collections.CollectionUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.*;

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
        HashMap<String, Object> excelMap = new HashMap<>();
        // 数据库操作
        List<Map<String, Object>> list = new ArrayList<>(); // odfSuggestionMapper.selectPoForecastList(rbc, plant);
        excelMap.put("datas", list);
        HashMap<String, Object> headerMap = new HashMap<>();
        headerMap.put("FCSTID", "FCSTID");
        headerMap.put("BOM", "BOM");
        headerMap.put("MODEL", "机型");
        headerMap.put("ITEM_DESC", "BOM描述");
        headerMap.put("FCSTDATE", "预测日期");
        headerMap.put("FCSTWEEK", "预测周次");
        headerMap.put("PLANT", "工厂");
        headerMap.put("ORG_ID", "RBC");
        headerMap.put("ORG_NAME", "RBC名称");
        headerMap.put("CUSTOMER", "客户编号");
        headerMap.put("QUANTITY", "原需求数量");
        headerMap.put("ODF_QTY", "已开单数量");
        headerMap.put("OPEN_QTY", "可开单数量");
        headerMap.put("SYS_UPDATE_BY", "预测人");
        excelMap.put("head", headerMap);
        byte[] bytes = ExcelUtils.exportExcel(excelMap);
        String date = DateUtils.formatDateToStr(new Date(), "yyyy-MM-dd");
        ExportVO vo = new ExportVO("FCST预测_" + date, bytes);
        // return ResponseUtils.setExcelResponse(response, vo); // 设置响应头
    }


}
