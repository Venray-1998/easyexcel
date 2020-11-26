package com.fwr.easyexcel;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.fwr.easyexcel.dao")
public class EasyexcelApplication {

    public static void main(String[] args) {
        SpringApplication.run(EasyexcelApplication.class, args);
    }

}
