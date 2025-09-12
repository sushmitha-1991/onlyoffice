/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.service.impl;

import java.io.*;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.*;

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.repository.DocTemplateRepository;
import com.dsp.OnlyOfficeBridge.utils.ThirdPartyDocGen;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import com.dsp.OnlyOfficeBridge.service.OnlyOfficeService;
import com.dsp.OnlyOfficeBridge.utils.Utils;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
/**
 * 
 */
@Service
public class OnlyOfficeServiceImpl implements OnlyOfficeService {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Value("${onlyoffice.url}")
	private String onlyOfficeUrl;
    @Value("${bulletin2.gen.template.id}")
    String bulletin2TemplateId;
    @Value("${lob.gen.template.id}")
    String lobTemplateId;
    @Value("${translation.gen.template.id}")
    String translationTemplateId;
    @Value("${verbatim.gen.template.id}")
    String verbatimTemplateId;
    @Value("${download.url}")
    String downloadKmUrl;
	@Autowired
	private Utils utils;
    @Autowired
    DocTemplateRepository docTemplateRepository;
    @Autowired
    ThirdPartyDocGen thirdPartyDocGen;

    @Value("${verbatim.gen.full.day.ls.template.id}")
    private String verbatimFullDayLSTemplateId;

    @Value("${verbatim.gen.full.day.rs.template.id}")
    private String verbatimFullDayRSTemplateId;

	@Override
	public Map<String, Object> getEditorConfig(Map<String, String> payload) {
		String step = payload.get("step");
        String applicationId = payload.get("applicationId");
        String fileName = payload.get("fileName");
		String existingDocumentKey = payload.get("documentKey");
		String existingDsp2Key = payload.get("dsp2Key"); // optional
		String editFlag = payload.get("edit"); // "true", "false", or null
        String emptyFlg =  payload.get("empty");
        String fileType = fileName.split("\\.")[1];
        System.out.println("fileType>>> " + fileType);
        String templateId = payload.get("templateId")==null?"0":payload.get("templateId");
		boolean isInitStep = "INIT".equalsIgnoreCase(step);
		String keyToUse;
		String downloadUrl;
		String title;
        String dsp2Key;
        if(existingDsp2Key!=null)
        {
            dsp2Key = existingDsp2Key;
            templateId ="0" ;
        }
        else
        {
            dsp2Key = UUID.randomUUID().toString();
        }

		if (isInitStep) {
			keyToUse = UUID.randomUUID().toString(); // new key
			downloadUrl = onlyOfficeUrl + "/only-office/download/1"; // static blank doc
			title = "Document.odt";
			System.out.println("INIT step: Generated new documentKey: " + keyToUse);
		} else {
			keyToUse = UUID.randomUUID().toString(); // new key even for edit
            if(emptyFlg.equals("1"))
            {
                downloadUrl = onlyOfficeUrl + "/only-office/download-empty-docx/"+fileName;
            }
            else
            {
                String refId = payload.get("refId")==null ? "0":payload.get("refId");
                downloadUrl = onlyOfficeUrl + "/only-office/download/" + dsp2Key + "/" + templateId + "/" + refId;
            }


            //download-empty-docx
			title = "Document-" + keyToUse + ".odt";
			System.out.println("EDIT step: Generated NEW documentKey for editing: " + keyToUse);
		}

		String callbackUrl = onlyOfficeUrl + "/only-office/callback?dsp2Key=" + dsp2Key + "&applicationId="+applicationId;

		// Determine if editing is allowed
		boolean canEdit = !"false".equalsIgnoreCase(editFlag); // null and "true" are treated as editable

		// Document config
		Map<String, Object> document = new HashMap<>();
//		document.put("fileType", fileType);
		document.put("key", keyToUse);
		document.put("title", title);
		document.put("url", downloadUrl);
		document.put("permissions", Map.of("edit", canEdit));

		// Editor config
		Map<String, Object> editorConfig = new HashMap<>();
		editorConfig.put("mode", "edit");
		editorConfig.put("callbackUrl", callbackUrl);
		editorConfig.put("coEditing", Map.of("mode", "fast", "change", true));
		editorConfig.put("customization", Map.of("forcesave", true, "save", true));

		// Final response
		Map<String, Object> response = new HashMap<>();
		response.put("document", document);
		response.put("documentType", "word");
		response.put("editorConfig", editorConfig);
		response.put("documentKey", keyToUse);
		response.put("dsp2Key", dsp2Key);

		return response;
	}

	@Override
    public ResponseEntity<Resource> downloadByKey(String key,String templateId,String refId) {
        String[] keys = key.split(",");

        try {
            if(Integer.parseInt(templateId)>0)
            {
                List<DocTemplate> docTemplates = docTemplateRepository.findByIdOrChieldIdOrderBySeq(Long.valueOf(templateId),Long.parseLong(templateId));
                if(templateId.equals(bulletin2TemplateId))
                {
                    return thirdPartyDocGen.genBulletin2Generation(key,docTemplates,refId);
                }
                if(templateId.equals(lobTemplateId))
                {
                    return thirdPartyDocGen.genLobGeneration(key,docTemplates,refId);
                }
                if(templateId.equals(translationTemplateId))
                {
                    return thirdPartyDocGen.genTranslationGeneration(key,docTemplates,refId);
                }
                if(templateId.equals(verbatimTemplateId))
                {
                    return thirdPartyDocGen.genVerbatimGeneration(key,docTemplates,refId);
                }
                if(templateId.equals(verbatimFullDayLSTemplateId))
                {
                    return thirdPartyDocGen.generateVerbatimFullDayLSFile(key,docTemplates,refId);
                }
                if(templateId.equals(verbatimFullDayRSTemplateId))
                {
                    return thirdPartyDocGen.generateVerbatimFullDayRSFile(key,docTemplates,refId);
                }
            }
            else
            {
                Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
                if (metadataOpt.isPresent()) {
                    ByteArrayResource resource = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
                    return ResponseEntity.ok()
                            .contentType(MediaType.APPLICATION_OCTET_STREAM)
                            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
                            .contentLength(resource.contentLength())
                            .body(resource);
                }
            }


        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }

        return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
    }
	
	
    public byte[] mergeDocxFiles(List<byte[]> docxFiles) throws Exception {
        XWPFDocument mergedDoc = new XWPFDocument();
        
        // ✅ Always set default margins (1 inch)
        utils.setDefaultMargins(mergedDoc);
        
        boolean firstFile = true;

        for (byte[] fileBytes : docxFiles) {
            try (ByteArrayInputStream bais = new ByteArrayInputStream(fileBytes);
                 XWPFDocument doc = new XWPFDocument(bais)) { 

                // Add page break before each file except the first
                if (!firstFile) {
                    XWPFParagraph breakPara = mergedDoc.createParagraph();
                    breakPara.setPageBreak(true);
                    mergedDoc.createParagraph(); // add empty paragraph for safety
                }
               
                firstFile = false;

                // Copy paragraphs with formatting
                for (XWPFParagraph p : doc.getParagraphs()) {
                	
                    XWPFParagraph newPara = mergedDoc.createParagraph();
                    newPara.setAlignment(p.getAlignment());
                    newPara.setSpacingAfter(p.getSpacingAfter());
                    newPara.setSpacingBefore(p.getSpacingBefore());
                    newPara.setStyle(p.getStyle());

                    for (XWPFRun run : p.getRuns()) {
                        XWPFRun newRun = newPara.createRun();

                        // Copy all text fragments safely
                        int textCount = run.getCTR().sizeOfTArray();
                        StringBuilder fullText = new StringBuilder();
                        for (int i = 0; i < textCount; i++) {
                            String text = run.getText(i);
                            if (text != null) {
                                fullText.append(text);
                            }
                        }
                        newRun.setText(fullText.toString());

                        // Copy formatting
                        newRun.setBold(run.isBold());
                        newRun.setItalic(run.isItalic());
                        newRun.setUnderline(run.getUnderline());
                        newRun.setColor(run.getColor());
                        newRun.setFontFamily(run.getFontFamily());
                        newRun.setFontSize(run.getFontSize());

                        if (run.getCTR().isSetRPr()) {
                            newRun.getCTR().setRPr(run.getCTR().getRPr());
                        }
                    }

                }

                // Copy tables
                for (XWPFTable table : doc.getTables()) {
                    XWPFTable newTable = mergedDoc.createTable();
                    newTable.getCTTbl().set(table.getCTTbl().copy());
                    newTable.setWidth("100%"); // fit to page width 
                }                
            }
        }

        // Write to byte[]
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        mergedDoc.write(baos);
        return baos.toByteArray();
    }

    //    public static byte[] mergeDocxFiles(List<byte[]> docxFiles) throws Exception {
//        if (docxFiles.isEmpty()) {
//            throw new IllegalArgumentException("No files provided for merging");
//        }
//
//        // Load the first document as the base
//        WordprocessingMLPackage target = WordprocessingMLPackage.load(new ByteArrayInputStream(docxFiles.get(0)));
//        MainDocumentPart targetPart = target.getMainDocumentPart();
//
//        // Append all remaining documents
//        for (int i = 1; i < docxFiles.size(); i++) {
//            WordprocessingMLPackage srcPkg = WordprocessingMLPackage.load(new ByteArrayInputStream(docxFiles.get(i)));
//            MainDocumentPart srcPart = srcPkg.getMainDocumentPart();
//
//            targetPart.getContent().addAll(srcPart.getContent());
//        }
//
//        // Save the merged DOCX
//        ByteArrayOutputStream baos = new ByteArrayOutputStream();
//        SaveToZipFile saver = new SaveToZipFile(target);
//        saver.save(baos);
//
//        return baos.toByteArray(); // ✅ this is a .docx byte[]
//    }
//    public static byte[] mergeDocxFiles(List<byte[]> docxFiles) throws IOException {
//        try (XWPFDocument mergedDoc = new XWPFDocument()) {
//            for (byte[] docxFile : docxFiles) {
//                try (XWPFDocument srcDoc = new XWPFDocument(new ByteArrayInputStream(docxFile))) {
//                    srcDoc.getParagraphs().forEach(p -> mergedDoc.createParagraph().createRun().setText(p.getText()));
//                    srcDoc.getTables().forEach(t -> mergedDoc.createTable().getCTTbl().set(t.getCTTbl()));
//                }
//            }
//
//            try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
//                mergedDoc.write(baos);
//                return baos.toByteArray();
//            }
//        }
//    }
//	public ResponseEntity<Resource> downloadByKey(String key) {
//		Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
//		if (metadataOpt.isPresent()) {
//			ByteArrayResource resource = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
//			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
//					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
//					.contentLength(resource.contentLength()).body(resource);
//		}
//		return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
//	}
    public ResponseEntity<Resource> downloadByKey1(String key) {
        try {
            // Split multiple keys by comma
            String[] keys = key.split(",");
            List<XWPFDocument> documents = new ArrayList<>();

            for (String singleKey : keys) {
                Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(singleKey.trim());
                if (metadataOpt.isPresent()) {
                    ByteArrayResource resource = utils.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

                    try (InputStream is = new ByteArrayInputStream(resource.getByteArray())) {
                        XWPFDocument doc = new XWPFDocument(is);
                        documents.add(doc);
                    }
                }
            }

            if (documents.isEmpty()) {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
            }

            // Merge documents
            XWPFDocument mergedDoc = new XWPFDocument();
            // ✅ Always set default margins (1 inch)
            utils.setDefaultMargins(mergedDoc);

            for (XWPFDocument doc : documents) {
                doc.getDocument().getBody().getTblArray(); // Force load
                doc.getDocument().getBody().getPArray();   // Force load
                mergedDoc.createParagraph(); // Page break between docs if needed
                mergedDoc.getDocument().getBody().addNewSectPr(); // section break
                mergedDoc.getDocument().getBody().addNewP().set( doc.getDocument().getBody().getPArray(0) );
               // mergedDoc.getDocument().getBody().addNewTbl().set( doc.getDocument().getBody().getTblArray(0) );
            }

            // Save merged doc to byte[]
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            mergedDoc.write(baos);

            ByteArrayResource mergedResource = new ByteArrayResource(baos.toByteArray());

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
                    .contentLength(mergedResource.contentLength())
                    .body(mergedResource);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }


	private String extractUuidFromUrl(String url) {
		return url.substring(url.lastIndexOf("/") + 1);
	}

	@Override
	public ResponseEntity<Map<String, Object>> handleCallback(String dsp2Key,String applicationId, Map<String, Object> payload) {
		try {
			int status = (int) payload.get("status");
			System.out.println("status>>>>>> " + status);

			if (status == 2 || status == 6) {
				String fileUrl = (String) payload.get("url");
				String key = (String) payload.get("key");
				System.out.println("url >>>>> " + fileUrl);
				System.out.println("key >>>>> " + key);
				System.out.println("dsp2Key >>>>>" + dsp2Key);

				// 1. Download file from OnlyOffice
				System.out.println(">>> Downloading file from fileUrl - " + fileUrl);
				File tempFile = File.createTempFile("edited-doc", ".docx");
				try (InputStream in = new URL(fileUrl).openStream();
						FileOutputStream out = new FileOutputStream(tempFile)) {
					in.transferTo(out);
				}

				System.out.println(">>> Calling Upload to OpenKm");
				String documentName = key + ".docx";

				// 2. Upload to OpenKM
				String openKmResponse = utils.uploadPdfToOpenKM(tempFile, documentName, "/okm:root/only-office/", "231", "51303",
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
					metadata.setOnlyOfficeUrl(fileUrl);
					metadata.setUpdatedBy(1L);
					metadata.setUpdatedAt(LocalDateTime.now());
					metadata.setDocumentKey(documentUuid);
                    metadata.setApplicationId(Long.parseLong(applicationId));
					System.out.println(">>> Updated existing OnlyOfficeMetadata record");
				} else {
					metadata = OnlyOfficeMetadata.builder().applicationId(Long.parseLong(applicationId)).documentKey(documentUuid).dsp2Key(dsp2Key)
							.downloadUrl(downloadUrl).onlyOfficeUrl(fileUrl).createdBy(1L).updatedBy(1L).build();
					System.out.println(">>> Created new OnlyOfficeMetadata record");
				}

				onlyOfficeMetadataRepository.save(metadata);

				// Clean up
				tempFile.delete();
			}

			return ResponseEntity.ok(Map.of("error", 0));
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(Map.of("error", 1));
		}
	}


}
