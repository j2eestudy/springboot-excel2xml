package com.yonyou.springboot.excel2xml.controller;

import com.yonyou.springboot.excel2xml.service.ExcelResolveService;
import com.yonyou.springboot.excel2xml.vo.ResultBean;
import com.yonyou.springboot.excel2xml.vo.ShowVO;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.List;


/**
 * @Author: shijq
 * @Date: 2019/3/6 10:07
 */
@RestController
@RequestMapping("/excel")
public class ExcelImportController {

    Logger log = LoggerFactory.getLogger(ExcelImportController.class);

    @Autowired
    private ExcelResolveService service;

    @Value("${file.uploadPath}")
    private String uploadPath;

    @GetMapping("/importExcel")
    public String importExcel2() {
        log.info("ssssssssssssssssssssssss");
        return "w22";
    }

    @PostMapping("/import")
    public ResultBean importExcel(@RequestParam("file") MultipartFile file){
        if(file != null){
            String fileName = file.getOriginalFilename();
            int size = (int) file.getSize();
            log.info(fileName + "-->" + size);
            try {
                List<ShowVO> showVOS = service.readExcell(file.getInputStream(), uploadPath);
                return new ResultBean(showVOS);
            } catch (IOException e) {
                e.printStackTrace();
                log.error(e.getMessage());
                return new ResultBean(e);
            }
        }else{
            return new ResultBean(new Exception("上传附件为空"));
        }
    }



}
