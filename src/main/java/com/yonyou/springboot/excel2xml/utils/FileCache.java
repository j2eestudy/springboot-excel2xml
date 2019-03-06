package com.yonyou.springboot.excel2xml.utils;

import java.util.HashMap;
import java.util.Map;

/**
 * @Author: shijq
 * @Date: 2019/3/6 18:52
 */
public class FileCache {

    private static Map<String,String> files = new HashMap<>();

    public static void set(String key,String value){
        files.put(key,value);
    }

    public static String get(String key){
        return files.get(key);
    }
}
