package com.zl.excelutils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.zl.excelutils.domain.AdsOcLackMatFillInfoPO;
import com.zl.excelutils.domain.ExcelVO;
import com.zl.excelutils.utils.ExcelUtils;
import org.apache.commons.collections.CollectionUtils;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@SpringBootTest
class ExcelUtilsApplicationTests {

    /**
     * 以下导入导出均配合PO对象
     * 自定义注解实现导入校验和导出数据的表头与顺序实现
     * 可以参考po对象中的注解
     **/

    @Test
    void testImport() {
        String filePath = "src/main/resources//excel/test.xlsx";
        // 获取Excel数据信息
        Map<String, Object> map = ExcelUtils.importExcelToList(filePath, AdsOcLackMatFillInfoPO.class);

        // 有报错直接返回提示信息
        if (map.get("error") != null) {
            List<String> error = JSONArray.parseArray(JSONObject.toJSONString(map.get("error"))).toJavaList(String.class);
            if (CollectionUtils.isNotEmpty(error)) {
                error.forEach(System.out::println);
            }
            return;
        }
        //  保存数据
        if (map.get("data") != null) {
            List<AdsOcLackMatFillInfoPO> data = JSONArray.parseArray(JSONObject.toJSONString(map.get("data")), AdsOcLackMatFillInfoPO.class);
            if (CollectionUtils.isNotEmpty(data)) {
                // 打开批处理 存储数据
                System.out.println("成功导入" + data.size() + "条");
                data.forEach(System.out::println);
            }
        }
    }


    @Test
    void testExport(){
        String fileName = "lack_mat_fill_info.xlsx";
        List<AdsOcLackMatFillInfoPO> pos = new ArrayList<>();
        for (int i = 0 ; i <= 3 ; i++){  // 构建数据
            AdsOcLackMatFillInfoPO po = new AdsOcLackMatFillInfoPO();
            po.setPdFactory("MASA");
            po.setOOutOrderCode("12345678");
            po.setMatName("欠料电子料");
            po.setLackMatNo("Z2C10101002882/ZXD1");
            po.setBulkLoadingDate("2021-09-08 08:49:19");
            po.setDelayDays(13);
            po.setFillWay("空运");
            po.setFillDate("2021-09-12");
            po.setFillSiNo("DMTK2102044");
            po.setFillTpNo("DOOS211279");
            po.setFillContainerNo("GCXU5822395");
            po.setFillDoneNum(994);
            po.setFillEtd("2021-03-17");
            po.setFillEta("2021-03-17");
            po.setFillAta("2021-03-17 18:12");
            po.setFillAtd("2021-03-14 19:02");
            pos.add(po);
        }
        // 直接放入数据集合
        ExcelVO<AdsOcLackMatFillInfoPO> excelVO = new ExcelVO<>(pos);
        byte[] bytes = ExcelUtils.exportExcel(excelVO);  // 返回字节流
        String filePath = "src/main/resources/excel/" + fileName;
        ExcelUtils.downloadExcel(excelVO, filePath);  // 输出到指定路径
    }


}
