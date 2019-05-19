package com.hyf.teseasypoi.service;

import com.hyf.teseasypoi.entity.Student;

import java.util.List;

/**
 * @author howinfun
 * @version 1.0
 * @desc
 * @date 2019/5/18
 * @company WCWC
 */
public interface StudentService {
    public List<Student> queryStudentByClazzId(Integer clazzId);
}
