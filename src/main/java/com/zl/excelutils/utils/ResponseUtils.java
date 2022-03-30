package com.zl.excelutils.utils;

import com.zl.excelutils.domain.ContentType;
import com.zl.excelutils.domain.ExportVO;
import com.zl.excelutils.exp.ResponseException;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.net.URLEncoder;

/**
 * @Description:
 * @Date: 2022/1/11 17:30
 */
public class ResponseUtils {
    public static byte[] setExcelResponse(HttpServletResponse response, ExportVO vo) {
        try {
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(vo.getFileName() + ".xlsx", "utf-8"));
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
        } catch (Exception e) {
            throw new ResponseException(-1, "设置响应对象异常");
        }
        return vo.getBytes();
    }

    /**
     * 根据文件类型自动设置响应头
     **/
    public static byte[] setFileResponse(HttpServletResponse response, ExportVO vo) {
        try {
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(vo.getFileName(), "utf-8"));
            ContentType type = ContentType.fromFileExtension(vo.getFileName());
            response.setContentType(String.format("%s;charset=utf-8", type != null ? type.getMimeType() : ""));
        } catch (Exception e) {
            throw new ResponseException(-1, "设置响应对象异常");
        }
        return vo.getBytes();
    }


    public static byte[] setFileResponse(HttpServletResponse response, ExportVO vo, String fileType) {
        try {
            response.setHeader("application/octet-stream", "attachment;filename=" + URLEncoder.encode(vo.getFileName(), "utf-8"));
            ContentType type = ContentType.fromFileExtension(fileType);
            response.setContentType(String.format("%s;charset=utf-8", type != null ? type.getMimeType() : ""));
            BufferedOutputStream out = new BufferedOutputStream(response.getOutputStream());
            out.write(vo.getBytes());
            out.flush();
            response.flushBuffer();
            out.close();
        } catch (Exception e) {
            throw new ResponseException(-1, "设置响应对象异常");
        }
        return vo.getBytes();
    }
}
