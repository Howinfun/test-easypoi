package com.hyf.teseasypoi.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.hyf.teseasypoi.entity.Clazz;
import com.hyf.teseasypoi.entity.Student;
import com.hyf.teseasypoi.service.ClazzService;
import com.hyf.teseasypoi.service.StudentService;
import com.hyf.teseasypoi.utils.ExcelUtils;
import net.sf.json.JSONObject;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
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

    @GetMapping("/export2")
    public void exportModel(HttpServletRequest req,HttpServletResponse resp){
        Workbook workbook=null;
        try {
            Workbook wb = null;
            // 读取模板文件路径
            File file = ResourceUtils.getFile("classpath:templates//siteTemplate.xlsx");
            String fileName = file.getName();
            FileInputStream fins = new FileInputStream(file);
            String[] fileSAtr = fileName.split("\\.");
            if (fileSAtr[1] != null) {
                // 若文件后缀名为xls则创建excel2003对象
                if (fileSAtr[1].equals("xls")) {
                    POIFSFileSystem fs = new POIFSFileSystem(fins);
                    wb = new HSSFWorkbook(fs);
                    workbook = new HSSFWorkbook();
                    // 若文件后缀名为xlsx则创建excel2007或以上版本对象
                } else if (fileSAtr[1].equals("xlsx")) {
                    wb = new XSSFWorkbook(fins);
                    workbook = new XSSFWorkbook();
                }

                // 设置边框样式
                CellStyle cellStyle = workbook.createCellStyle();
                // 下边框
                cellStyle.setBorderBottom(BorderStyle.THIN);
                // 左边框
                cellStyle.setBorderLeft(BorderStyle.THIN);
                // 上边框
                cellStyle.setBorderTop(BorderStyle.THIN);
                // 右边框
                cellStyle.setBorderRight(BorderStyle.THIN);
                // 垂直居中
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                // 设置字体内容值水平居中
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setWrapText(true);
                // 模板页
                Sheet sheetModel = null;
                // 新建的Sheet页
                Sheet newSheet = null;
                // 创建行
                Row row = null;
                // 创建列
                Cell cell = null;
                // 循环建立Sheet页（结算明细）
                for (int i=1;i<=2000;i++) {
                    List list1 = new ArrayList<>();
                    Map<String,Object> listMap1 = new HashMap<>();
                    listMap1.put("dbcode","{{!fe:siteList"+i+" t.dbcode");
                    listMap1.put("bl","t.bl");
                    listMap1.put("j","t.j");
                    listMap1.put("f","t.f");
                    listMap1.put("p","t.p");
                    listMap1.put("g","t.g");
                    listMap1.put("dl","t.dl");
                    listMap1.put("fee","t.fee}}");
                    /*String[] strArr = new String[10];
                    strArr[0]="{{!fe:siteList"+i+"t.dbcode";
                    strArr[0]="t.bl";
                    strArr[0]="t.j";
                    strArr[0]="t.f";
                    strArr[0]="t.p";
                    strArr[0]="t.g";
                    strArr[0]="t.dl";
                    strArr[0]="t.fee}}";*/
                    list1.add(listMap1);

                    // 读取模板中模板Sheet页中的内容
                    sheetModel = wb.getSheetAt(0);
                    // 设置新建Sheet的页名
                    newSheet = workbook.createSheet("站点"+i);
                    // 将模板中的内容复制到新建的Sheet页中
                    copySheet(workbook, sheetModel, newSheet, sheetModel.getFirstRowNum(), sheetModel.getLastRowNum(),cellStyle);
                    //获取到新建Sheet页中的第二行为其中的列赋值
                    row = newSheet.getRow(1);
                    row.getCell(1).setCellValue("{{site"+i+".siteName}}");
                    //注意 合并的单元格也要按照合并前的格数来算
                    row.getCell(5).setCellValue("{{site"+i+".settlementTime}}");

                    //获取第3行
                    row = newSheet.getRow(2);
                    row.getCell(2).setCellValue("{{site"+i+".monthdl}}");
                    row.getCell(5).setCellValue("{{site"+i+".electricityFee}}");
                    //获取第4行
                    row = newSheet.getRow(3);
                    row.getCell(0).setCellValue("{{site"+i+".rule}}");
                    row.getCell(2).setCellValue("{{site"+i+".monthcddl}}");
                    row.getCell(5).setCellValue("{{site"+i+".parkingFeeDate}}");
                    //获取第5行
                    row = newSheet.getRow(4);
                    row.getCell(2).setCellValue("{{site"+i+".parkingNum}}");
                    row.getCell(5).setCellValue("{{site"+i+".fee}}");
                    //获取第8行
                    row = newSheet.getRow(7);
                    row.getCell(1).setCellValue("{{site"+i+".Proprietor}}");
                    row.getCell(4).setCellValue("{{site"+i+".Operator}}");
                    //获取第9行
                    row = newSheet.getRow(8);
                    row.getCell(1).setCellValue("{{site"+i+".ProprietorTrustees}}");
                    row.getCell(4).setCellValue("{{site"+i+".OperatorTrustees}}");
                    //获取第10行
                    row = newSheet.getRow(9);
                    row.getCell(4).setCellValue("{{site"+i+".currentDay}}");

                    // 设置第一行表头标题单元格
                    String[] str = {"电表编号","CT变比","尖时电价","峰时电价","平时电价","谷时电价","电表电量","电表费用"};
                    String[] fieldArr1 = { "dbcode","bl","j","f","p","g","dl","fee"};
                    writeInSheet(workbook,list1, str, fieldArr1, newSheet, cellStyle, 12);
                }
            }
            fins.close();
        }catch (Exception e){

        }
        downLoadExcel(resp,workbook,"xxx.xlsx");
    }

    private void copySheet(Workbook wb, Sheet fromsheet, Sheet newSheet, int firstrow, int lasttrow, CellStyle cellStyle) {

        // 复制一个单元格样式到新建单元格
        if ((firstrow == -1) || (lasttrow == -1) || lasttrow < firstrow) {
            return;
        }
        // 复制合并的单元格
        CellRangeAddress mergedRegion = null;
        for (int i = 0; i < fromsheet.getNumMergedRegions(); i++) {
            mergedRegion = fromsheet.getMergedRegion(i);
            if ((mergedRegion.getFirstRow() >= firstrow) && (mergedRegion.getLastRow() <= lasttrow)) {
                newSheet.addMergedRegion(mergedRegion);
            }
        }
        Row fromRow = null;
        Row newRow = null;
        Cell newCell = null;
        Cell fromCell = null;
        // 设置列宽
        for (int i = firstrow; i < lasttrow; i++) {
            fromRow = fromsheet.getRow(i);
            if (fromRow != null) {
                for (int j = fromRow.getLastCellNum(); j >= fromRow.getFirstCellNum(); j--) {
                    int colnum = fromsheet.getColumnWidth((short) j);
                    if (colnum > 100) {
                        newSheet.setColumnWidth((short) j, (short) colnum);
                    }
                    if (colnum == 0) {
                        newSheet.setColumnHidden((short) j, true);
                    } else {
                        newSheet.setColumnHidden((short) j, false);
                    }
                }
                break;
            }
        }

        // 复制行并填充数据
        for (int i = 0; i <=lasttrow; i++) {
            fromRow = fromsheet.getRow(i);
            if (fromRow == null) {
                continue;
            }
            newRow = newSheet.createRow(i - firstrow);
            newRow.setHeight(fromRow.getHeight());
            for (int j = fromRow.getFirstCellNum(); j <=fromRow.getPhysicalNumberOfCells(); j++) {
                fromCell = fromRow.getCell((short) j);
                if (fromCell == null) {
                    continue;
                }
                newCell = newRow.createCell((short) j);
                /*poi按照一个源单元格设置目标单元格格式，如果两个单元格不在同一个workbook，
                  要用 HSSFCellStyle下的cloneStyleFrom()方法，而不能用HSSFCell下的setCellStyle()方法。*/
                //设置样式
                newCell.setCellStyle(cellStyle);
                //单元格的类型
                int cType = fromCell.getCellType();
                newCell.setCellType(cType);
                switch (cType) {
                    case Cell.CELL_TYPE_STRING:
                        newCell.setCellValue(fromCell.getRichStringCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        newCell.setCellValue(fromCell.getNumericCellValue());
                        break;
                    //公式型，一定要用setCellValue才能设置成功
                    case Cell.CELL_TYPE_FORMULA:
                        newCell.setCellFormula(fromCell.getCellFormula());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        newCell.setCellValue(fromCell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        newCell.setCellValue(fromCell.getErrorCellValue());
                        break;
                    default:
                        newCell.setCellValue(fromCell.getRichStringCellValue());
                        break;
                }

            }
        }
    }

    private Sheet writeInSheet(Workbook workbook,List<Object[]> list,  String[] tableHead, String[] fieldArr,Sheet sheet1,CellStyle style,int index) {

        // 创建行
        // 第一行,表头
        // TODO: 2018/7/30
        Font font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short)9);//设置字体大小
        style.setFont(font);
        Row row1 = sheet1.createRow(index);
        row1.setHeightInPoints(30);
        // row1.setRowStyle(style);
        // 设置表头行数据
        for (int i = 0; i < tableHead.length; i++) {
            // 设置每列单元格的宽度为23个字符宽度；i表示第几列
            sheet1.setColumnWidth(i, 20 * 256);
            Cell headCell = row1.createCell(i);
            headCell.setCellStyle(style);
            headCell.setCellValue(tableHead[i]);
        }

        // 循环创建数据行
        for (int i = 0; i <list.size(); i++) {
            Row row = sheet1.createRow(i+index+1);
            row.setHeightInPoints(30);
            Object obj = list.get(i);
            JSONObject jsonObj = JSONObject.fromObject(obj);
            for (int j = 0; j < fieldArr.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(style);
                if (StringUtils.isBlank(jsonObj.get(fieldArr[j]) + "")
                        || "null".equals(jsonObj.get(fieldArr[j]) + "")) {
                    cell.setCellValue("");
                } else {
                    cell.setCellValue(jsonObj.get(fieldArr[j]) + "");
                }
            }
        }
        return sheet1;
    }

    private void downLoadExcel(HttpServletResponse resp, Workbook wb, String fileName) {
        // ServletOutputStream outSTr;
        resp.reset();
        resp.setCharacterEncoding("GBK");
        // 设置文件类型
        resp.setContentType("APPLICATION/XLS");
        // 设置头部信息
        try (ServletOutputStream outSTr = resp.getOutputStream()) {
            resp.setHeader("Content-Disposition",
                    "attachment; filename=" + new String(fileName.getBytes("gb2312"), "ISO8859-1"));
            // outSTr = resp.getOutputStream();
            // 设置第一行表头标题单元格
            wb.write(outSTr);
            outSTr.flush();

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
