package com.fwr.easyexcel.entity;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

import java.io.Serializable;
import java.util.Date;

/**
 * Student
 * 
 * @author  fwr
 * @date 2020-11-17 
 */
@Data
@TableName("boss")
public class Boss implements Serializable {

	@TableId(type = IdType.AUTO)
	private Integer id;

	private String name;

	private Date birthday;

	private Double score;

}
