package com.yonyou.springboot.excel2xml.vo;

import lombok.Data;

import java.io.Serializable;

/**
 * @Author: shijq
 * @Date: 2019/3/6 19:23
 */
@Data
public class ResultBean<T> implements Serializable {

    private static final long serialVersionUID = 2646904085153784595L;
    private static final int NO_LOGIN = -1;
    private static final int SUCCESS = 0;
    private static final int FAIL = 1;
    private static final int NO_PERMISSION = 2;
    private String msg = "success";
    private int code = SUCCESS;
    private T data;
    public ResultBean(){
        super();
    }
    public ResultBean(T data){
        super();
        this.data = data;
    }
    public ResultBean(Throwable e){
        super();
        this.msg = e.getMessage();
        this.code = FAIL;
    }
}
