package com.hyf.teseasypoi.service.impl;

import com.hyf.teseasypoi.entity.Student;
import com.hyf.teseasypoi.mapper.StudentMapper;
import com.hyf.teseasypoi.service.StudentService;
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
public class StudentServiceImpl implements StudentService {

    @Autowired
    private StudentMapper studentMapper;

    public List<Student> queryStudentByClazzId(Integer clazzId){
        Student student = new Student();
        student.setClazzId(clazzId);
        return studentMapper.select(student);
    }
}
