package com.hyf.teseasypoi.service.impl;

import com.hyf.teseasypoi.entity.Clazz;
import com.hyf.teseasypoi.mapper.ClazzMapper;
import com.hyf.teseasypoi.service.ClazzService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @author howinfun
 * @version 1.0
 * @desc
 * @date 2019/5/18
 * @company WCWC
 */
@Service
public class ClazzServiceImpl implements ClazzService {

    @Autowired
    private ClazzMapper clazzMapper;

    public List<Clazz> queryClazzList(){
        return clazzMapper.selectAll();
    }
}
