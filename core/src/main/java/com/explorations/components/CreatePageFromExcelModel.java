package com.explorations.components;

import com.explorations.services.CreatePageFromExcelService;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.models.annotations.Model;

import javax.annotation.PostConstruct;
import javax.inject.Inject;

@Model(adaptables = SlingHttpServletRequest.class)
public class CreatePageFromExcelModel {

    @Inject
    private CreatePageFromExcelService createPageFromExcelService;

    @Inject
    private SlingHttpServletRequest request;

    private String excelServiceResult;

    @PostConstruct
    public void init() {
        excelServiceResult = createPageFromExcelService.GetExcelWorkbook(request);
    }

    public String getExcelServiceResult() {
        return excelServiceResult;
    }

}
