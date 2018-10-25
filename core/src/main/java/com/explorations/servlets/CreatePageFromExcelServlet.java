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

    private static final String SUCCESS_MESSAGE = "Page/s created successfully from excel.";

    private static final String FORMAT_FAILURE_MESSAGE = "Format not supported, Please try again with a proper .xls file";

    private static final String PAGE_CREATION_FAILURE_MESSAGE = "Error occurred while trying to create pages using the excel sheet";

    private static final String NOT_MULTIPART_FAILURE_MESSAGE = "The content uploaded is not a multipart content";

    private static final String NO_FILE_FAILURE_MESSAGE = "No proper file is uploaded, please upload an .xls file to proceed";

    private static final Logger LOGGER = LoggerFactory.getLogger(CreatePageFromExcelServlet.class);

    @Override
    protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response) throws IOException {
        if (FileUpload.isMultipartContent(new ServletRequestContext(request))) {
            RequestParameter param = request.getRequestParameter(UPLOADED_FILE);
            if (param != null && param.getSize() >0) {
                createPageFromExcel(request, param.getInputStream(), response);
            } else response.getWriter().append(NO_FILE_FAILURE_MESSAGE).flush();
        } else response.getWriter().append(NOT_MULTIPART_FAILURE_MESSAGE).flush();
    }

    private void createPageFromExcel(SlingHttpServletRequest request, InputStream inputStream, SlingHttpServletResponse response) throws IOException {
        try {
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
                response.getWriter().append(SUCCESS_MESSAGE).flush();
            }
        } catch (BiffException e) {
            LOGGER.error("Encountered a Biff exception while using extracting a workbook", e);
            response.getWriter().append(FORMAT_FAILURE_MESSAGE).flush();
        } catch (WCMException e) {
            LOGGER.error("Encountered an exception while creating a page using PageManager", e);
            response.getWriter().append(PAGE_CREATION_FAILURE_MESSAGE).flush();
        }
    }
}
