/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.service.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

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

/**
 * 
 */
@Service
public class OnlyOfficeServiceImpl implements OnlyOfficeService {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Value("${onlyoffice.url}")
	private String onlyOfficeUrl;

	@Autowired
	private Utils utils;

	@Override
	public Map<String, Object> getEditorConfig(Map<String, String> payload) {
		String step = payload.get("step");
        String fileName = payload.get("fileName");
		String existingDocumentKey = payload.get("documentKey");
		String existingDsp2Key = payload.get("dsp2Key"); // optional
		String editFlag = payload.get("edit"); // "true", "false", or null
        String emptyFlg =  payload.get("empty");
        String fileType = fileName.split("\\.")[1];
        System.out.println("fileType>>> " + fileType);
		boolean isInitStep = "INIT".equalsIgnoreCase(step);
		String keyToUse;
		String downloadUrl;
		String title;
		String dsp2Key = (existingDsp2Key != null) ? existingDsp2Key : UUID.randomUUID().toString();

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
                downloadUrl = onlyOfficeUrl + "/only-office/download/" + dsp2Key;
            }


            //download-empty-docx
			title = "Document-" + keyToUse + ".odt";
			System.out.println("EDIT step: Generated NEW documentKey for editing: " + keyToUse);
		}

		String callbackUrl = onlyOfficeUrl + "/only-office/callback?dsp2Key=" + dsp2Key;

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
	public ResponseEntity<Resource> downloadByKey(String key) {
		Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
		if (metadataOpt.isPresent()) {
			String latestUuid = extractUuidFromUrl(metadataOpt.get().getDownloadUrl());
			ByteArrayResource resource = utils.downloadFileFromUUID(UUID.fromString(latestUuid));
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
					.contentLength(resource.contentLength()).body(resource);
		}
		return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
	}



	private String extractUuidFromUrl(String url) {
		return url.substring(url.lastIndexOf("/") + 1);
	}

	@Override
	public ResponseEntity<Map<String, Object>> handleCallback(String dsp2Key, Map<String, Object> payload) {
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
				String downloadUrl = "https://dspstg.sansad.in/dms/openkm/dsp1/downloads/" + documentUuid;

				// 5. Update or insert into OnlyOfficeMetadata
				Optional<OnlyOfficeMetadata> existingMetadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dsp2Key);

				OnlyOfficeMetadata metadata;
				if (existingMetadataOpt.isPresent()) {
					metadata = existingMetadataOpt.get();
					metadata.setDownloadUrl(downloadUrl);
					metadata.setOnlyOfficeUrl(fileUrl);
					metadata.setUpdatedBy(1L);
					metadata.setUpdatedAt(LocalDateTime.now());
					metadata.setDocumentKey(key);
					System.out.println(">>> Updated existing OnlyOfficeMetadata record");
				} else {
					metadata = OnlyOfficeMetadata.builder().applicationId(117L).documentKey(key).dsp2Key(dsp2Key)
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
