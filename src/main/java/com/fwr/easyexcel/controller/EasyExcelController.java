package com.fwr.easyexcel.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.fwr.easyexcel.dao.BossMapper;
import com.fwr.easyexcel.entity.Boss;
import com.fwr.easyexcel.entity.BossReadExcel;
import com.fwr.easyexcel.entity.BossWriteExcel;
import com.fwr.easyexcel.service.EasyExcelService;
import com.fwr.easyexcel.util.ExcelUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import sun.java2d.loops.FillParallelogram;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author fwr
 * @date 2020-11-24
 */
@Api(tags = "首页模块")
@RestController
@RequestMapping("")
public class EasyExcelController {

    @Autowired
    private EasyExcelService easyExcelService;
    @Autowired
    private BossMapper bossMapper;

    @ApiOperation(value = "向客人问好")
    @PostMapping("/upload")
    public void upload(MultipartFile file) throws IOException {
        List<BossReadExcel> bossReadExcels = ExcelUtil.readExcel(file, BossReadExcel.class);
        List<Boss> bosses = bossReadExcels.stream().map(bossReadExcel -> {
            Boss boss = new Boss();
            BeanUtils.copyProperties(bossReadExcel, boss);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
            Date parse = null;
            try {
                parse = simpleDateFormat.parse(bossReadExcel.getBirthday());
            } catch (ParseException e) {
                e.printStackTrace();
            }
            boss.setBirthday(parse);
            return boss;
        }).collect(Collectors.toList());

        int i = bossMapper.insertList(bosses);
        System.out.println(i);
    }

    @GetMapping("/download")
    public void download(HttpServletResponse response) throws IOException {
        List<Boss> bosses = bossMapper.selectList(null);
        List<BossWriteExcel> bossWriteExcels = bosses.stream().map(boss -> {
            BossWriteExcel bossWriteExcel = new BossWriteExcel();
            BeanUtils.copyProperties(boss, bossWriteExcel);
            return bossWriteExcel;
        }).collect(Collectors.toList());
        ExcelUtil.writeExcel(response, bossWriteExcels, BossWriteExcel.class, "boss", "boss");
    }

    /**
     * 多张sheet
     *
     * @param response
     * @throws IOException
     */
    @GetMapping("/download2")
    public void download2(HttpServletResponse response) throws IOException {
        List<Boss> bosses = bossMapper.selectList(null);
        List<BossWriteExcel> bossWriteExcels = bosses.stream().map(boss -> {
            BossWriteExcel bossWriteExcel = new BossWriteExcel();
            BeanUtils.copyProperties(boss, bossWriteExcel);
            return bossWriteExcel;
        }).collect(Collectors.toList());

        String fileName = "test";
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), BossWriteExcel.class).build();
        WriteSheet writeSheet1 = EasyExcel.writerSheet(0,"test1").build();
        WriteSheet writeSheet2 = EasyExcel.writerSheet(1,"test2").build();
        excelWriter.write(bossWriteExcels, writeSheet1);
        excelWriter.write(bossWriteExcels, writeSheet2);
        excelWriter.finish();
    }

    @GetMapping("/downloadByTemplate")
    public void downloadByTemplate(HttpServletResponse response) throws IOException {
        List<Boss> bosses = bossMapper.selectList(null);
        List<BossWriteExcel> bossWriteExcels = bosses.stream().map(boss -> {
            BossWriteExcel bossWriteExcel = new BossWriteExcel();
            BeanUtils.copyProperties(boss, bossWriteExcel);
            return bossWriteExcel;
        }).collect(Collectors.toList());
        String templateName = this.getClass().getClassLoader().getResource("./template3.xlsx").getPath();
        //ExcelUtil.writeExcel(response, bossWriteExcels, BossWriteExcel.class, templateName,"boss", "boss");

        HashMap<String, String> map = new HashMap<>();
        map.put("date","1998");
        map.put("total", "10");

        ExcelWriter excelWriter = EasyExcel.write("hhh.xlsx",BossWriteExcel.class)
                .withTemplate(templateName)
                .build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();

        //水平填充数据
        FillConfig fillConfig2 = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        excelWriter.fill(new FillWrapper("data1", bossWriteExcels), fillConfig2, writeSheet);

        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        FillConfig fillConfig1 = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        // 用fillwrapper区分每个数据
        excelWriter.fill(new FillWrapper("data2", bossWriteExcels), fillConfig1, writeSheet);
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();
        System.out.println(1);
    }

    /**
     * 浏览器下载 多个sheet
     *
     * @param response
     * @throws IOException
     */
    @GetMapping("/downloadByTemplate2")
    public void downloadByTemplate2(HttpServletResponse response) throws IOException {
        List<Boss> bosses = bossMapper.selectList(null);
        List<BossWriteExcel> bossWriteExcels = bosses.stream().map(boss -> {
            BossWriteExcel bossWriteExcel = new BossWriteExcel();
            BeanUtils.copyProperties(boss, bossWriteExcel);
            return bossWriteExcel;
        }).collect(Collectors.toList());
        String templateName = this.getClass().getClassLoader().getResource("./template4.xlsx").getPath();
        HashMap<String, String> map = new HashMap<>();
        map.put("date","1998");
        map.put("total", "10");

        String fileName = "test";
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");

        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), BossWriteExcel.class)
                .withTemplate(templateName)
                .build();
        //模板填充的sheetName是跟模板的，下面的test1，test不生效
        WriteSheet writeSheet1 = EasyExcel.writerSheet(0,"test1").build();
        WriteSheet writeSheet2 = EasyExcel.writerSheet(1,"test2").build();
        //水平填充数据
        FillConfig fillConfig2 = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        excelWriter.fill(new FillWrapper("data1", bossWriteExcels), fillConfig2, writeSheet1);

        // 这里注意 入参用了forceNewRow 代表在写入list的时候不管list下面有没有空行 都会创建一行，然后下面的数据往后移动。默认 是false，会直接使用下一行，如果没有则创建。
        // forceNewRow 如果设置了true,有个缺点 就是他会把所有的数据都放到内存了，所以慎用
        FillConfig fillConfig1 = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        // 用fillwrapper区分每个数据
        excelWriter.fill(new FillWrapper("data2", bossWriteExcels), fillConfig1, writeSheet1);
        excelWriter.fill(map, writeSheet1);

        excelWriter.fill(new FillWrapper("data1", bossWriteExcels), fillConfig2, writeSheet2);
        excelWriter.fill(new FillWrapper("data2", bossWriteExcels), fillConfig1, writeSheet2);
        excelWriter.fill(map, writeSheet2);

        excelWriter.finish();
    }

}
