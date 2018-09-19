package com;

import org.springframework.web.bind.annotation.RequestMapping;

public class HelloSpringMvcWord {

    private static final String SUCCESS = "success"; //定义常量

    /**
     * 1.使用@RequestMapping 注解来映射请求的url
     * 2.返回值会通过视图解析为实际的物理视图 ，对于InternalResourceViewResolver 视图解析，会做如下解析
     *  通过prefix + returnVal + 后缀得到实际的物理视图，然后做转发操作
     *  /WEB-INF/views/success.jsp
     * @return
     */
    @RequestMapping("/helloword")
    public String hello(){
        System.out.println("Spring MVC word;you are ok...");
        return SUCCESS;
    }
}
