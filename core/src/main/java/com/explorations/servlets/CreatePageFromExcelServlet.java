package com.explorations.servlets;

import com.day.cq.wcm.api.WCMException;
import com.explorations.services.CreatePageFromExcelService;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.apache.sling.api.servlets.SlingSafeMethodsServlet;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.Servlet;
import java.io.IOException;

@Component(
        service = Servlet.class,
        property = {
                "sling.servlet.methods=POST",
                "sling.servlet.resourceTypes=sling/servlet/default",
                "sling.servlet.selectors=createPageFromExcel",
                "sling.servlet.extensions=json",
                "service.description=explorations: Create page from excel servlet"
        }
)
public class CreatePageFromExcelServlet extends SlingAllMethodsServlet {

    private static final String SELECTED_EXCEL = "selectedExcel";

    private static final Logger LOGGER = LoggerFactory.getLogger(CreatePageFromExcelServlet.class);

    @Reference
    private CreatePageFromExcelService createPageFromExcelService;

    @Override
    protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response) throws IOException {
        try {
            String selectedExcelResourcePath = request.getParameter(SELECTED_EXCEL);
            if (createPageFromExcelService.createPageFromExcel(request, selectedExcelResourcePath)) {
                response.getWriter().append("Page/Pages created successfully!").flush();
            }
            response.getWriter().append("Failed creating pages from excel").flush();

        } catch (WCMException e) {
            LOGGER.error("An Exception occurred while trying to create pages from excel", e);
            response.getWriter().print("Error occurred while creating pages from excel");
        }
    }

}
