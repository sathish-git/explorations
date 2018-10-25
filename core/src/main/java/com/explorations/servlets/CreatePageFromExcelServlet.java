package com.explorations.servlets;

import com.day.cq.wcm.api.PageManager;
import com.day.cq.wcm.api.WCMException;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.commons.fileupload.FileUpload;
import org.apache.commons.fileupload.servlet.ServletRequestContext;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.request.RequestParameter;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.osgi.service.component.annotations.Component;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.Servlet;
import java.io.IOException;
import java.io.InputStream;

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

    private static final String UPLOADED_FILE = "uploadedFile";

    private static final Logger LOGGER = LoggerFactory.getLogger(CreatePageFromExcelServlet.class);

    @Override
    protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response) throws IOException {
        try {
            boolean isCreatePageFromExcelSuccess;
            if (FileUpload.isMultipartContent(new ServletRequestContext(request))) {
                RequestParameter param = request.getRequestParameter(UPLOADED_FILE);
                isCreatePageFromExcelSuccess = param != null && createPageFromExcel(request, param.getInputStream());
                if (isCreatePageFromExcelSuccess) {
                    response.getWriter().append("Page/s successfully created from excel!").flush();
                } else response.getWriter().append("Cannot create pages from excel, please try again").flush();
            } else response.getWriter().append("! The form is not a multipart form.");
        } catch (NullPointerException | WCMException | BiffException e) {
            LOGGER.error("An Exception occurred while trying to create pages from excel", e);
            response.getWriter().print("Error occurred while creating pages from excel");
        }
    }

    private boolean createPageFromExcel(SlingHttpServletRequest request, InputStream inputStream) throws WCMException, IOException, BiffException {
        Workbook workbook = inputStream != null ? Workbook.getWorkbook(inputStream) : null;
        PageManager pageManager = request.getResourceResolver().adaptTo(PageManager.class);
        if (workbook != null && pageManager != null) {
            Sheet sheet = workbook.getSheet(0);
            for (int i = 1; i < sheet.getRows(); i++) {
                pageManager.create(
                        sheet.getCell(0, i).getContents(),
                        sheet.getCell(1, i).getContents(),
                        sheet.getCell(2, i).getContents(),
                        sheet.getCell(3, i).getContents(), true);
            }
            return true;
        }
        return false;
    }
}
