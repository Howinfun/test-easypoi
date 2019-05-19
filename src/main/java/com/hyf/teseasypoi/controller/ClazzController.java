package com.hyf.teseasypoi.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.hyf.teseasypoi.entity.Clazz;
import com.hyf.teseasypoi.entity.Student;
import com.hyf.teseasypoi.service.ClazzService;
import com.hyf.teseasypoi.service.StudentService;
import com.hyf.teseasypoi.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author howinfun
 * @version 1.0
 * @desc
 * @date 2019/5/18
 * @company WCWC
 */
@RestController
@RequestMapping("/clazz")
public class ClazzController {

    @Autowired
    private ClazzService clazzService;
    @Autowired
    private StudentService studentService;

    @GetMapping("/export")
    public void export(HttpServletResponse response){
        // 模版数据
        Map<String,Object> map = new HashMap<String,Object>() ;
        List<Clazz> clazzList = clazzService.queryClazzList();
        Integer count = 0;
        map.put("clazzList",clazzList);
        String[] sheetName = new String[clazzList.size()+1];
        sheetName[0] = "班级信息";
        for (Clazz clazz : clazzList) {
            count++;
            sheetName[count] = clazz.getName()+"统计信息";
            Integer clazzId = clazz.getId();
            List<Student> students = studentService.queryStudentByClazzId(clazzId);
            Map<String,Object> clazzMap = new HashMap<>();
            clazzMap.put("name",clazz.getName());
            clazzMap.put("count",students.size());
            map.put("clazz"+count,clazzMap);
            map.put("stuList"+count,students);

        }
        // 设置导出配置
        // 获取导出excel指定模版
        TemplateExportParams params = new TemplateExportParams("templates/Clazz3.xlsx",false);
        // 导出的sheet
        Integer[] sheetNum = new Integer[clazzList.size()+1];
        for (int i=0;i<=clazzList.size();i++){
            sheetNum[i]=i;
        }
        params.setSheetNum(sheetNum);
        // 设置sheetName(多个sheet)，若不设置该参数，则使用得原本得sheet名称
        params.setSheetName(sheetName);
        Workbook workbook = ExcelExportUtil.exportExcel(params,map);
        ExcelUtils.downLoadExcel("xxx.xlsx",response,workbook);
    }
}
