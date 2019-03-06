package com.yonyou.springboot.excel2xml.controller;

import com.yonyou.springboot.excel2xml.utils.FileCache;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

/**
 * @Author: shijq
 * @Date: 2019/3/6 19:29
 */
@Controller
public class XmlDownloadController {

    @GetMapping("filedownload")
    public void downloadXML(@RequestParam("uuid") String uuid, HttpServletRequest req, HttpServletResponse response) {

        String path = FileCache.get(uuid);

        File file = new File(path);

        response.setContentType("application/force-download");// 设置强制下载不打开
        try {
            response.setHeader("Content-Disposition", "attachment; fileName="+  file.getName() +";filename*=utf-8''"+ URLEncoder.encode(file.getName(),"UTF-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        byte[] buffer = new byte[1024];
        FileInputStream fis = null;
        BufferedInputStream bis = null;
        try {
            fis = new FileInputStream(file);
            bis = new BufferedInputStream(fis);
            OutputStream os = response.getOutputStream();
            int i = bis.read(buffer);
            while (i != -1) {
                os.write(buffer, 0, i);
                i = bis.read(buffer);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bis != null) {
                try {
                    bis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
