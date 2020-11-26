package com.fwr.easyexcel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import lombok.Data;

import java.util.Date;

/**
 * @author fwr
 * @date 2020-11-25
 */
@Data
public class BossWriteExcel {
    @ExcelProperty(value = "老板id",index = 0)
    private Integer id;

    @ExcelProperty(value = "老板名称",index = 1)
    private String name;

    @ExcelProperty(value = "老板生日",index = 2)
    @DateTimeFormat("yyyy-MM-dd")
    private Date birthday;

    @ExcelProperty(value = "老板分数", index = 3)
    @NumberFormat("#.####")
    private Double score;
}
