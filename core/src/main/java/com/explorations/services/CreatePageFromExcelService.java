package com.explorations.services;

import com.day.cq.dam.api.Asset;
import jxl.Workbook;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.osgi.service.component.annotations.Component;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Iterator;

@Component(service = CreatePageFromExcelService.class)
public class CreatePageFromExcelService {

    private final Logger LOGGER = LoggerFactory.getLogger(CreatePageFromExcelService.class);
    private static final String EXCEL_ASSET_PATH = "/content/dam/explorations/excel";
    private static final String EXCEL_DC_FORMAT = "application/vnd.ms-excel";


    public String GetExcelWorkbook(SlingHttpServletRequest request) {
        try {
            ResourceResolver resourceResolver = request.getResourceResolver();
            Resource folderResource = resourceResolver.getResource(EXCEL_ASSET_PATH);
            if (folderResource != null) {
                Iterator<Resource> resourceIterator = folderResource.listChildren();
                while (resourceIterator.hasNext()) {
                    Resource resource = resourceIterator.next();
                    Asset asset = resource.adaptTo(Asset.class);
                    if (asset != null && asset.getMimeType().equals(EXCEL_DC_FORMAT)) {
                        return "excel found!!";
                    }

                }
            }
        } catch (Exception e) {
            LOGGER.error("An Exception has occurred", e);
        }
        return "No excel found";
    }

}
