package com.explorations.beans;

import org.apache.commons.lang3.StringUtils;

public class CreatePageFromExcelBean {

    private String excelName;
    private String excelResourcePath;

    public CreatePageFromExcelBean(String excelName, String excelResourcePath) {
        this.excelName = excelName;
        this.excelResourcePath = excelResourcePath;
    }

    public String getExcelName() {
        return StringUtils.substringBefore(excelName, ".").trim();
    }

    public String getExcelResourcePath() {
        return excelResourcePath;
    }

}
