package com.zl.excelutils.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @Description: 合并单元格实体类
 * @Date: 2021/10/27 10:31
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class MergeExcelVO {

    private String[] head; // 合并单元格内容 example: String[] head0 = new String[]{"卡板数:", 12, "物料款数:", 15};

    private String[] mergeNum; // 合并单元格信息  example: String[] headnum0 = new String[]{"0,0,0,1", "0,0,2,2", "0,0,3,4", "0,0,5,5"};


}
