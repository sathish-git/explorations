package com.explorations.components;

import com.explorations.beans.CreatePageFromExcelBean;
import com.explorations.services.CreatePageFromExcelService;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.models.annotations.Model;

import javax.annotation.PostConstruct;
import javax.inject.Inject;
import java.util.List;

@Model(adaptables = SlingHttpServletRequest.class)
public class CreatePageFromExcelModel {

    @Inject
    private CreatePageFromExcelService createPageFromExcelService;

    @Inject
    private SlingHttpServletRequest request;

    private List<CreatePageFromExcelBean> excelsAvailable;

    @PostConstruct
    public void init() {
        excelsAvailable = createPageFromExcelService.getAllExcels(request);
    }

    public List<CreatePageFromExcelBean> getExcelsAvailable() {
        return excelsAvailable;
    }
}
