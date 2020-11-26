package com.fwr.easyexcel.dao;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.fwr.easyexcel.entity.Boss;
import org.apache.ibatis.annotations.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * @author fwr
 * @date 2020-11-24
 */
@Repository
public interface BossMapper extends BaseMapper<Boss> {

    int insertList(@Param("list") List<Boss> list);
}
