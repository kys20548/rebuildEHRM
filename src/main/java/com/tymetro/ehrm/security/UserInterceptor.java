/*
 *
 *
 * 之前是用這隻，之後改sporing security


package com.tymetro.ehrm.security;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.servlet.handler.HandlerInterceptorAdapter;

public class UserInterceptor extends HandlerInterceptorAdapter {

    @Value("${user_name}")
    public String USE_RNAME;
    @Value("${user_password}")
    public String USER_PASSWORD;

    @Override
    public boolean preHandle(HttpServletRequest request, HttpServletResponse response, Object handler)
            throws Exception {
        String session_userName=(String)request.getSession().getAttribute("tymetro_ehrm_user_name");
        String session_userPassword=(String)request.getSession().getAttribute("tymetro_ehrm_user_password");
        if(isNotEmptyString(session_userName)&&isNotEmptyString(session_userPassword)){
            if(session_userName.equals(USE_RNAME)&&session_userPassword.equals(USER_PASSWORD)) {
                return true;
            }else {
                response.sendRedirect("/ehrm/login.jsp");
                return false;
            }
        }else {
            response.sendRedirect("/ehrm/login.jsp");
            return false;
        }
    }

    public boolean isNotEmptyString(String obj) {
        if(obj==null) {
            return false;
        }else if(obj.trim().equals("")) {
            return false;
        }else {
            return true;
        }
    }
}
*/
