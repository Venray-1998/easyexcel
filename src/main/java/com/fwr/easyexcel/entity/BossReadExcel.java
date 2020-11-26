package com.fwr.easyexcel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

/**
 * @author fwr
 * @date 2020-11-24
 */
@Data
public class BossReadExcel {

    @ExcelProperty(value = "老板id",index = 0)
    private Integer id;

    @ExcelProperty(value = "老板名称",index = 1)
    private String name;

    @ExcelProperty(value = "老板生日",index = 2)
    @DateTimeFormat("yyyy-MM-dd")
    private String birthday;

    @ExcelProperty(value = "老板分数",index = 3)
    private Double score;
}
