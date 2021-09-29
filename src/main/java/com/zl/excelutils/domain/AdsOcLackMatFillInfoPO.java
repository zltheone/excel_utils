package com.zl.excelutils.domain;

import com.zl.excelutils.annotations.*;
import lombok.Data;

@Data
@ExcelSheetName("欠料补料表")
public class AdsOcLackMatFillInfoPO {

    @ExcelImportIgnoreField
    private String id;

    @ExcelExportHeadName("生产工厂")
    private String pdFactory;

    @ExcelExportHeadName("订单号")
    @ExcelImportNotNull
    @ExcelJointValueSingle("订单号,欠料物编,补料TP单")
    private String outOrderCode;

    @ExcelExportHeadName("物料名称")
    private String matName;

    @ExcelExportHeadName("欠料物编")
    @ExcelImportNotNull
    @ExcelJointValueSingle("订单号,欠料物编,补料TP单")
    private String lackMatNo;

    @ExcelExportHeadName("大货装柜日期")
    @ExcelRequestDate
    private String bulkLoadingDate;

    @ExcelExportHeadName("大货装柜欠料数")
    private Integer bulkLoadingLackQty;

    @ExcelExportHeadName("延误天数")
    private Integer delayDays;

    @ExcelExportHeadName("补料方式")
    @ExcelAssignValue("空运,夹柜,快递")
    private String fillWay;

    @ExcelExportHeadName("补料时间")
    @ExcelRequestDate
    private String fillDate;

    @ExcelExportHeadName("补料SI号")
    private String fillSiNo;

    @ExcelExportHeadName("补料TP单")
    @ExcelImportNotNull
    @ExcelJointValueSingle("订单号,欠料物编,补料TP单")
    private String fillTpNo;

    @ExcelExportHeadName("补料柜号")
    private String fillContainerNo;

    @ExcelExportHeadName("已补数量")
    private Integer fillDoneNum;

    @ExcelExportHeadName("补料ETD")
    @ExcelRequestDate
    private String fillEtd;

    @ExcelExportHeadName("补料ATD")
    @ExcelRequestDate("yyyy-MM-dd HH:mm")
    private String fillAtd;

    @ExcelExportHeadName("补料ETA")
    @ExcelRequestDate
    private String fillEta;

    @ExcelExportHeadName("补料ATA")
    @ExcelRequestDate("yyyy-MM-dd HH:mm")
    private String fillAta;

    @ExcelExportHeadName("备注")
    private String note;

    @ExcelImportIgnoreField
    private String outOrderCodeHzerp;

    @ExcelImportIgnoreField
    private String oOutOrderCode;

    @ExcelImportIgnoreField
    private String oOutOrderCodeHzerp;

    @ExcelImportIgnoreField
    private String sSt;

    @ExcelImportIgnoreField
    private Integer isFillDone;

    @ExcelImportIgnoreField
    private String excelTpId;

    @ExcelImportIgnoreField
    private String operator;

}