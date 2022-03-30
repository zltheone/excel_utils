package com.zl.excelutils.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExportVO {

    private String fileName;

    private byte[] bytes;

    private String fileType;

    public ExportVO(String fileName, byte[] bytes) {
        this.fileName = fileName;
        this.bytes = bytes;
    }
}
