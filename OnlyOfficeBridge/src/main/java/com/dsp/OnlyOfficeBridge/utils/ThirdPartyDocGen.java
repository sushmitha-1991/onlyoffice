package com.dsp.OnlyOfficeBridge.utils;

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.*;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestTemplate;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

@Component
public class ThirdPartyDocGen {
    @Autowired
    private Utils utils;
    @Autowired
    private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;
//    @Autowired
//    RestTemplate restTemplate;
    @Value("${bulletin2.service.url}")
    private String bulletin2Url;
    public ResponseEntity<Resource> genBulletin2Generation(List<DocTemplate> docTemplates, String refId)
    {
        String url = bulletin2Url + "/get-generate-doc/"+refId;
        System.out.println("url>>>>>>>" + url);
        String token = "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ0ZXN0dXNlcjEiLCJob3VzZUlkIjoyLCJob3VzZUFiYnIiOiJMUyIsImlwIjoiMTAuMjYuMTA1LjEyMCIsImlhdCI6MTc1NjY0MzY4MywiZXhwIjoxNzU2NzMwMDgzfQ.cU-sLJg-2_zid9TnJX6LkUsvFPoYcq5-LFawhtGKwNZ6NWZllIuo2VTtpzpw-y_MEzbPy_YzzVKvHtW3j9RjiLku6z5iIyLeBlICFtBfiCqxTKhcYSvw72gGdGqw3xxHwqg6JD3RluZ9xZOyAeeSX6zTcfQFaqLP8knuDy4313Yt1kx9v89dhh1qBYQpqCirVGloRy35iQePOr3YTu1jh1Q5-fjVFMGWk3UYt0WLlzPnkTha0Da1vrPRO3tmfeUtTJfi10Mz4YoTCJbbzP_gLM-PK25IqL1vP34JZbIN8fyCtmtfXDO_Zyl9GJIeNFfeIFWonsuy-f8kXW15gdSsDA";//utils.getTokenFromRequest();
        Map<String,Object> body;
//        if (token == null) {
//            throw new IllegalStateException("Authorization token is missing");
//        }
        HttpHeaders headers = new HttpHeaders();
        headers.set("Authorization", "Bearer " + token);
        headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
        HttpEntity<Void> entity = new HttpEntity<>( headers);

        try {
            RestTemplate restTemplate = new RestTemplate();
            ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
             body = response.getBody();
            System.out.println(body);
        } catch (Exception e) {
            System.err.println("Error during API call: " + e.getMessage());
            throw e;
        }

        List<byte[]> files = new LinkedList<>();

        docTemplates.forEach(docTemplate -> {
            System.out.println("header uuid >>>>" + docTemplate.getDocUuid());

            if(docTemplate.getCode().equals("bull2_gen_header")) {
                Map<String,String> map = new HashMap<>();
                Map<String,Object> data = (Map<String, Object>) body.get("data");
                map.put("range",(Integer)data.get("startNum") + " - "  + (Integer) data.get("endNum"));
                map.put("date",(String)data.get("forDate") + "/" + (String)data.get("forDateSaka"));
                try {
                    ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);

                    files.add(resource.getByteArray());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }

            }
            else if (docTemplate.getCode().equals("bull2_gen_con")) {
                Map<String,Object> data = (Map<String, Object>) body.get("data");
                List<Map> inputList = (List<Map>)data.get("inputList");
                inputList.forEach(map -> {
                    Map<String,String> map1 = new HashMap<>();
                    map1.put("myno", String.valueOf((Integer)map.get("bulletinNo")));
                    map1.put("branch",  (String)map.get("officeName"));
                    map1.put("subject",  (String)map.get("subject"));

                    ByteArrayResource resource = null;
                    try {
                        resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map1);
                    } catch (Exception e) {
//                        e.printStackTrace();
                        throw new RuntimeException(e);
                    }
                    files.add(resource.getByteArray());
                    Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key((String)map.get("documentKey"));
                    ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
                    files.add(resourceBody.getByteArray());
                });

            }
            else {
                ByteArrayResource resource = null ;
                try {
                    Map map = new HashMap();
                    resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
                files.add(resource.getByteArray());
            }
        });
        try {
            byte[] mergedBytes = mergeDocxFiles(files);
            ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
                    .contentLength(mergedResource.contentLength())
                    .body(mergedResource);
        } catch (Exception e) {

        }
        return null;
    }

    public static byte[] mergeDocxFiles(List<byte[]> docxFiles) throws Exception {
        XWPFDocument mergedDoc = new XWPFDocument();

        for (byte[] fileBytes : docxFiles) {
            try (ByteArrayInputStream bais = new ByteArrayInputStream(fileBytes);
                 XWPFDocument doc = new XWPFDocument(bais)) {

                CTBody srcBody = doc.getDocument().getBody();
                CTBody destBody = mergedDoc.getDocument().getBody();

                appendBody(destBody, srcBody);
            }
        }

        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            mergedDoc.write(baos);
            return baos.toByteArray();
        }
    }

    private static void appendBody(CTBody dest, CTBody src) throws Exception {
        XmlCursor cursor = dest.newCursor();
        cursor.toEndToken(); // move to end
        src.newCursor().copyXml(cursor); // append source body elements
        cursor.dispose();
    }


}
