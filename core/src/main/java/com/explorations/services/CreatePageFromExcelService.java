package com.explorations.services;

import com.day.cq.dam.api.Asset;
import com.day.cq.wcm.api.PageManager;
import com.day.cq.wcm.api.WCMException;
import com.explorations.beans.CreatePageFromExcelBean;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.osgi.service.component.annotations.Component;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

@Component(service = CreatePageFromExcelService.class)
public class CreatePageFromExcelService {

    private static final String EXCEL_ASSET_PATH = "/content/dam/explorations/excel";
    private static final String EXCEL_DC_FORMAT = "application/vnd.ms-excel";
    private final Logger LOGGER = LoggerFactory.getLogger(CreatePageFromExcelService.class);

    public List<CreatePageFromExcelBean> getAllExcels(SlingHttpServletRequest request) {
        try {
            ResourceResolver resourceResolver = request.getResourceResolver();
            Resource folderResource = resourceResolver.getResource(EXCEL_ASSET_PATH);
            List<CreatePageFromExcelBean> excelsList = new ArrayList<>();
            if (folderResource != null) {
                Iterator<Resource> resourceIterator = folderResource.listChildren();
                while (resourceIterator.hasNext()) {
                    Resource resource = resourceIterator.next();
                    Asset asset = resource.adaptTo(Asset.class);
                    if (asset != null && asset.getMimeType().equals(EXCEL_DC_FORMAT)) {
                        excelsList.add(new CreatePageFromExcelBean(asset.getName(), asset.getPath()));
                    }
                }
                return excelsList;
            }
        } catch (Exception e) {
            LOGGER.error("An Exception has occurred while retrieving the excels in DAM", e);
        }
        return Collections.emptyList();
    }

    public boolean createPageFromExcel(SlingHttpServletRequest request, String excelResourcePath) throws WCMException {
        Workbook workbook = getWorkbook(request, excelResourcePath);
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

    public Workbook getWorkbook(SlingHttpServletRequest request, String excelResourcePath) {
        try {
            Resource excelResource = request.getResourceResolver().getResource(excelResourcePath);
            if (excelResource != null) {
                Asset excelAsset = excelResource.adaptTo(Asset.class);
                if (excelAsset != null && !excelAsset.getRenditions().isEmpty()) {
                    InputStream inputStream = excelAsset.getRendition("original").getStream();
                    return Workbook.getWorkbook(inputStream);
                }
            }
        } catch (Exception e) {
            LOGGER.error("An Exception has occurred while retrieving jxl workbook for selected excel", e);
        }
        return null;
    }

}
