package com.dsp.OnlyOfficeBridge.utils;

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
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

import java.io.*;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

@Slf4j
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
    @Value("${lob.service.url}")
    private String lobUrl;
    @Value("${admin.token}")
    String token;
    @Value("${download.url}")
    String downloadKmUrl;
    public ResponseEntity<Resource> genBulletin2Generation(String dsp2Key,List<DocTemplate> docTemplates, String refId)
    {
        String url = bulletin2Url + "/get-generate-doc/"+refId;
        System.out.println("url>>>>>>>" + url);
        Map<String,Object> body;
        if (token == null) {
            throw new IllegalStateException("Authorization token is missing");
        }
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
                    Map<String,String> mapLine = new HashMap<>();
                    ByteArrayResource resource = null;
                    ByteArrayResource resourceLine = null;
                    try {
                        resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map1);
                        resourceLine = Utils.replacePlaceholders("line.docx", mapLine);


                    } catch (Exception e) {
//                        e.printStackTrace();
                        throw new RuntimeException(e);
                    }
                    files.add(resource.getByteArray());
                    Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key((String)map.get("dsp2Key"));
                    ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
                    files.add(resourceBody.getByteArray());
                    files.add(resourceLine.getByteArray());
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
            saveOpenKm(dsp2Key,mergedResource);
            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
                    .contentLength(mergedResource.contentLength())
                    .body(mergedResource);
        } catch (Exception e) {
            log.info(e.getMessage());
        }
        return null;
    }
    public ResponseEntity<Resource> genLobGeneration(String dsp2Key,List<DocTemplate> docTemplates, String refId)
    {
        String url = lobUrl + "/lob/generate-doc-id/"+refId;
        System.out.println("url>>>>>>>" + url +" ==== " + token);
        Map<String,Object> body;
        if (token == null) {
            throw new IllegalStateException("Authorization token is missing");
        }
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

            if(docTemplate.getCode().equals("lob_gen_header")) {
                Map<String,String> map = new HashMap<>();
                Map<String,Object> data = (Map<String, Object>) body.get("data");
                map.put("date",(String)data.get("forDate") + "/" + (String)data.get("forDateSaka"));
                try {
                    ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);

                    files.add(resource.getByteArray());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }

            }
            else if (docTemplate.getCode().equals("lob_gen_con")) {
                Map<String,Object> data = (Map<String, Object>) body.get("data");
                List<Map> inputList = (List<Map>)data.get("inputList");
                inputList.forEach(map -> {
                    Map<String,String> map1 = new HashMap<>();
                    if(map.get("dsp2Key")!=null) {
                        map1.put("subject", String.valueOf(map.get("subjectStr")));

                        Map<String, String> mapLine = new HashMap<>();
                        ByteArrayResource resource = null;
                        ByteArrayResource resourceLine = null;
                        try {
                            resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map1);
                            resourceLine = Utils.replacePlaceholders("line.docx", mapLine);

                        } catch (Exception e) {
//                        e.printStackTrace();
                            throw new RuntimeException(e);
                        }
                        files.add(resource.getByteArray());
                        Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key((String) map.get("dsp2Key"));
                        ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
                        files.add(resourceBody.getByteArray());
                        files.add(resourceLine.getByteArray());
                    }
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
            saveOpenKm(dsp2Key,mergedResource);
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
    public void saveOpenKm(String dsp2Key,ByteArrayResource resource)
    {
        try{
            File tempFile = File.createTempFile("edited-doc", ".docx");

            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                fos.write(resource.getByteArray());
            }
            System.out.println(">>> Calling Upload to OpenKm");
            String documentName = dsp2Key + ".docx";

            // 2. Upload to OpenKM
            String openKmResponse = utils.uploadPdfToOpenKM(tempFile, documentName, "/okm:root/", "231", "51303",
                    "ApplicationInstance", "1");

            System.out.println(">>> Download completed" + openKmResponse);

            // 3. Parse OpenKM response
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(openKmResponse);
            JsonNode dataNode = rootNode.path("data");
            String documentUuid = dataNode.path("documentUuid").asText();

            System.out.println(">>> documentUuid - " + documentUuid);

            // 4. Build OpenKM download URL
            String downloadUrl = downloadKmUrl + documentUuid;

            // 5. Update or insert into OnlyOfficeMetadata
            Optional<OnlyOfficeMetadata> existingMetadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dsp2Key);

            OnlyOfficeMetadata metadata;
            if (existingMetadataOpt.isPresent()) {
                metadata = existingMetadataOpt.get();
                metadata.setDownloadUrl(downloadUrl);
                metadata.setOnlyOfficeUrl(downloadUrl);
                metadata.setUpdatedBy(1L);
                metadata.setUpdatedAt(LocalDateTime.now());
                metadata.setDocumentKey(documentUuid);
                System.out.println(">>> Updated existing OnlyOfficeMetadata record");
            } else {
                metadata = OnlyOfficeMetadata.builder().applicationId(0L).documentKey(documentUuid).dsp2Key(dsp2Key)
                        .downloadUrl(downloadUrl).onlyOfficeUrl(downloadUrl).createdBy(1L).updatedBy(1L).build();
                System.out.println(">>> Created new OnlyOfficeMetadata record");
            }

            onlyOfficeMetadataRepository.save(metadata);

            // Clean up
            tempFile.delete();
        }
        catch (Exception e)
        {
            log.info(e.getMessage());
        }

    }

}
