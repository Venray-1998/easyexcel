package com.fwr.easyexcel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.fastjson.JSON;
import com.fwr.easyexcel.entity.BossWriteExcel;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 这里只是简单的读写，其余具体用法参考 https://www.yuque.com/easyexcel
 *
 * @author fwr
 * @date 2020-11-24
 */
public class ExcelUtil {

    /**
     * 读取全部sheet
     *
     * @param file
     * @param rowModel
     * @return java.util.List<org.apache.poi.ss.formula.functions.T>
     */
    public static<T> List<T> readExcel(MultipartFile file, Class<T> rowModel) throws IOException {
        checkExcelType(file);
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 这里需要注意 DemoDataListener的doAfterAllAnalysed 会在每个sheet读取完毕后调用一次。然后所有sheet都会往同一个DemoDataListener里面写,读取完会自动关闭
        ExcelListener<T> excelListener = new ExcelListener<>();
        EasyExcel.read(file.getInputStream(), rowModel, excelListener)
                .doReadAll();
        return excelListener.getList();
    }

    /**
     * 读取指定sheet
     *
     * @param file
     * @param rowModel
     * @param sheetNo
     * @return
     * @throws IOException
     */
    public static<T> List<T> readExcel(MultipartFile file, Class<T> rowModel, int sheetNo) throws IOException {
        checkExcelType(file);
        ExcelListener<T> excelListener = new ExcelListener<>();
        EasyExcel.read(file.getInputStream(), rowModel, excelListener)
                .sheet(sheetNo)
                .doRead();
        return excelListener.getList();
    }

    /**
     * 导出excel,存在一张sheet
     *
     * @param response
     * @param data
     * @param rowModel
     * @param fileName
     * @param sheetName
     * @throws IOException
     */
    public static<T> void writeExcel(HttpServletResponse response, List<T> data, Class<T> rowModel, String fileName, String sheetName) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), rowModel)
                    .excelType(ExcelTypeEnum.XLSX)
                    .autoCloseStream(Boolean.FALSE)
                    .sheet(sheetName)
                    .doWrite(data);
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    /**
     * 根据模板填充导出excel,存在一张sheet
     *
     * @param response
     * @param data
     * @param rowModel
     * @param fileName
     * @param sheetName
     * @throws IOException
     */
    public static<T> void writeExcel(HttpServletResponse response, List<T> data, Class<T> rowModel, String templatePath,String fileName, String sheetName) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            // 这里需要设置不关闭流
//            ExcelWriterBuilder excelWriterBuilder = EasyExcel.write(response.getOutputStream(), rowModel).withTemplate(templatePath).autoCloseStream(Boolean.FALSE);
//            ExcelWriterSheetBuilder sheet = excelWriterBuilder.sheet();
//            sheet.doFill(data);
            HashMap<String, String> map = new HashMap<>();
            map.put("date","1998");
            map.put("total", "10");
            ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), rowModel)
                    .withTemplate(templatePath)
                    .build();
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
            // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
            excelWriter.fill(data, fillConfig, writeSheet);
            excelWriter.fill(map, writeSheet);
            excelWriter.finish();
        } catch (Exception e) {
            // 重置response
            e.printStackTrace();
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    private static void checkExcelType(MultipartFile file) {
        String excelType = file.getOriginalFilename();
        if (excelType == null || (!excelType.toLowerCase().endsWith(".xls") && !excelType.toLowerCase().endsWith(".xlsx"))) {
            throw new RuntimeException("文件格式错误！");
        }
    }
}