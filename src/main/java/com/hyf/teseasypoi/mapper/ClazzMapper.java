package com.hyf.teseasypoi.mapper;

import com.hyf.teseasypoi.entity.Clazz;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Component;
import tk.mybatis.mapper.common.BaseMapper;

/**
 * @author howinfun
 * @version 1.0
 * @desc
 * @date 2019/5/18
 * @company WCWC
 */
@Mapper
@Component
public interface ClazzMapper extends BaseMapper<Clazz> {
}
