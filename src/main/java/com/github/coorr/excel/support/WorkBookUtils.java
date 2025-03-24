package com.github.coorr.excel.support;

import org.springframework.util.StringUtils;

public abstract class WorkBookUtils {
    public static String createHeaderValue(String headerName, String secondHeaderName) {
        if (StringUtils.hasText(headerName)) {
            if (StringUtils.hasText(secondHeaderName)) {
                return headerName + "-" + secondHeaderName;
            }
            return headerName;
        }
        return secondHeaderName;
    }
}
