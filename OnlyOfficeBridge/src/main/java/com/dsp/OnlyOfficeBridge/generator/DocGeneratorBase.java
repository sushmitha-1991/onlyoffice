package com.dsp.OnlyOfficeBridge.generator;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.RestTemplate;

import com.dsp.OnlyOfficeBridge.dto.MargeFileDto;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public abstract class DocGeneratorBase {

	@Value("${admin.token}")
	String token;
	
	@Value("${dsp2.base.url}")
	String baseUrl;
	
	@Value("${download.url}")
	String downloadOkmUrl;
	
	@Value("${upload.url}")
	String uploadOkmUrl; 
	
	@Autowired
	private RestTemplate restTemplate;
	
	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;
	
	abstract String getBaseUrl(String refId);
	
	String url;
	
	/**
	 * 
	 */
	public HttpEntity getEntityForRestCall() {
		
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		
		log.debug("token>>>>>>> {}", token);

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);
		return entity;
	}
	
	
	Map<String, Object> getResponse () {
		
		HttpEntity entity = getEntityForRestCall();
		
		Map<String, Object> body;

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			log.debug("Response Body: {}", body);
		} catch (Exception e) {
			log.error("Error during API call-B2: {}" + e.getMessage());
			throw e;
		}
		
		return body;
	}
	
	/**
	 * 
	 * @param files
	 * @param dsp2Key
	 * @return
	 */
	ResponseEntity mergeFiles (List<MargeFileDto> files, String dsp2Key) {
		
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {

		}
		return null;
	}
	
    /**
     * 
     * @param dsp2Key
     * @param resource
     */
	public void saveOpenKm(String dsp2Key, ByteArrayResource resource) {
		try {
			File tempFile = File.createTempFile("edited-doc", ".docx");

			try (FileOutputStream fos = new FileOutputStream(tempFile)) {
				fos.write(resource.getByteArray());
			}
			log.debug(">>> Calling Upload to OpenKm");
			String documentName = dsp2Key + ".docx";

			// 2. Upload to OpenKM
			String openKmResponse = uploadPdfToOpenKM(tempFile, documentName, "/okm:root/", "231", "51303",
					"ApplicationInstance", "1");

			log.debug(">>> Download completed {}", openKmResponse);

			// 3. Parse OpenKM response
			ObjectMapper objectMapper = new ObjectMapper();
			JsonNode rootNode = objectMapper.readTree(openKmResponse);
			JsonNode dataNode = rootNode.path("data");
			String documentUuid = dataNode.path("documentUuid").asText();

			log.debug(">>> documentUuid - ", documentUuid);

			// 4. Build OpenKM download URL
			String downloadUrl = downloadOkmUrl + documentUuid;

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
				log.debug(">>> Updated existing OnlyOfficeMetadata record");
			} else {
				metadata = OnlyOfficeMetadata.builder().applicationId(0L).documentKey(documentUuid).dsp2Key(dsp2Key)
						.downloadUrl(downloadUrl).onlyOfficeUrl(downloadUrl).createdBy(1L).updatedBy(1L).build();
				log.debug(">>> Created new OnlyOfficeMetadata record");
			}

			onlyOfficeMetadataRepository.save(metadata);

			// Clean up
			tempFile.delete();
		} catch (Exception e) {
			log.info(e.getMessage());
		}
	}	
	
	/**
	 * 
	 * @param pdfFile
	 * @param documentName
	 * @param folderPath
	 * @param masterId
	 * @param referenceId
	 * @param referenceType
	 * @param applicationNo
	 * @return
	 */
	public String uploadPdfToOpenKM(File pdfFile, String documentName, String folderPath, String masterId,
			String referenceId, String referenceType, String applicationNo) {
		try {
			System.out.println(">>> calling uploadPdfToOpenKM -");
			System.out.println(">>> documentName - " + documentName);
			System.out.println(">>> folderPath - " + folderPath);
			System.out.println(">>> masterId - " + masterId);
			System.out.println(">>> referenceId - " + referenceId);
			System.out.println(">>> referenceType - " + referenceType);
			System.out.println(">>> applicationNo - " + applicationNo);

			// The JWT token is automatically added by the RestTemplate interceptor
			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.MULTIPART_FORM_DATA);

			// Create a FileSystemResource from the file
			FileSystemResource fileResource = new FileSystemResource(pdfFile);

			// Prepare the body with all required fields
			MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
			body.add("folderPath", folderPath);
			body.add("documentName", documentName);
			body.add("file", fileResource);
			body.add("masterId", masterId);
			body.add("referenceId", referenceId);
			body.add("referenceType", referenceType);
			body.add("applicationNo", applicationNo);

			System.out.println(">>> inside");

			// Prepare the request entity with headers and body
			HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);
			System.out.println(requestEntity);
//			String uploadUrl = "https://dspstg.sansad.in/dms/openkm/dsp1/document/upload";
			System.out.println(">>> NEW uploadUrl - " + uploadOkmUrl); 
			ResponseEntity<String> response = restTemplate.exchange(uploadOkmUrl, HttpMethod.POST, requestEntity,
					String.class);

			// Return the response body as a string
			return response.getBody();
		} catch (HttpClientErrorException e) {
			e.printStackTrace();
			throw new RuntimeException("Error uploading document to OpenKM. HTTP Error: " + e.getMessage());
		} catch (Exception e) {
			throw new RuntimeException("Error uploading document to OpenKM.", e);
		}
	}
		
	/**
	 * 
	 * @param uuid
	 * @return
	 */
	public ByteArrayResource downloadFileFromUUID(UUID uuid) {
		try {
			String downloadUrl = downloadOkmUrl + uuid;
			System.out.println(">>>>> DownloadURL: " + downloadUrl);
			
			RestTemplate restTemplate = new RestTemplate();

			HttpHeaders headers = new HttpHeaders();
			HttpEntity<Void> entity = new HttpEntity<>(headers);

			ResponseEntity<byte[]> response = restTemplate.exchange(downloadUrl, HttpMethod.GET, entity, byte[].class);

			if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
				String filename = uuid.toString(); // Default filename

				List<String> contentDisposition = response.getHeaders().get(HttpHeaders.CONTENT_DISPOSITION);
				if (contentDisposition != null && !contentDisposition.isEmpty()) {
					// Example: Content-Disposition: attachment; filename="Document.docx"
					String disposition = contentDisposition.get(0);
					int start = disposition.indexOf("filename=");
					if (start != -1) {
						filename = disposition.substring(start + 9).replace("\"", "").trim();
					}
				}

				final String finalFilename = filename;
				return new ByteArrayResource(response.getBody()) {
					@Override
					public String getFilename() {
						return finalFilename;
					}
				};
			}
		} catch (Exception e) {
			e.printStackTrace(); // Use proper logging
		}

		return null;
	}
	
	   /**
     * 
     * @param resourcePath
     * @param replacements
     * @return
     * @throws Exception
     */
    public ByteArrayResource replacePlaceholders(String resourcePath, Map<String, String> replacements) throws Exception {

        // Load file from resources
        ClassPathResource resource = new ClassPathResource("doc-template/" + resourcePath);

        // Copy resource to a temp file (LibreOffice requires a real File, not just InputStream)
        File inputFile = File.createTempFile("template", ".docx");
        try (InputStream in = resource.getInputStream();
             OutputStream out = new FileOutputStream(inputFile)) {
            in.transferTo(out);
        }

        File tempFile = File.createTempFile("converted", ".fodt");

        LocalOfficeManager officeManager = LocalOfficeManager.install();
        officeManager.start();
//        RemoteOfficeManager officeManager = RemoteOfficeManager.builder()
//                // If your LibreOffice server runs in socket mode:
//                //.urlConnection("socket://<REMOTE_SERVER_IP>:2002")
//
//                // If your LibreOffice server exposes REST API:
//                .urlConnection("http://10.247.128.22:2002")
//                .build();

        try {
            // Convert DOCX → FODT
            LocalConverter.builder()
                    .officeManager(officeManager)
                    .build()
                    .convert(inputFile)
                    .to(tempFile)
                    .as(DefaultDocumentFormatRegistry.FODT)
                    .execute();

            // Read + replace placeholders in FODT XML
            String xml = new String(java.nio.file.Files.readAllBytes(tempFile.toPath()), StandardCharsets.UTF_8);

            for (Map.Entry<String, String> entry : replacements.entrySet()) {
//                String placeholder = "\\{" + entry.getKey() + "\\}";
//                xml = xml.replaceAll(placeholder, entry.getValue());
                String placeholder = "{" + entry.getKey() + "}";
                String safeValue = StringEscapeUtils.escapeXml11(entry.getValue()); // escapes &, <, >, ", '
                xml = xml.replace(placeholder, safeValue);
            }

            // Write modified FODT back
            java.nio.file.Files.write(tempFile.toPath(), xml.getBytes(StandardCharsets.UTF_8));

            // Convert FODT → DOCX into memory
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            LocalConverter.builder()
                    .officeManager(officeManager)
                    .build()
                    .convert(tempFile)
                    .to(baos)
                    .as(DefaultDocumentFormatRegistry.DOCX)
                    .execute();

            return new ByteArrayResource(baos.toByteArray());

        } finally {
            officeManager.stop();
            inputFile.delete(); // cleanup temp files
            tempFile.delete();
        }
    }
    
	public static byte[] mergeDocxFiles(List<MargeFileDto> margeFileDtos) throws Exception {
		log.info("Total files to merge: " + margeFileDtos.size());

		XWPFDocument mergedDoc = new XWPFDocument();

		for (int i = 0; i < margeFileDtos.size(); i++) {
			MargeFileDto fileBytes = margeFileDtos.get(i);
			try (ByteArrayInputStream bais = new ByteArrayInputStream(fileBytes.getFile());
					XWPFDocument doc = new XWPFDocument(bais)) {

//				// ---- Copy body paragraphs ----
//				for (XWPFParagraph p : doc.getParagraphs()) {
//					XWPFParagraph newP = mergedDoc.createParagraph();
//					newP.getCTP().set(p.getCTP().copy());
//				}
//
//				// ---- Copy tables ----
//				for (XWPFTable table : doc.getTables()) {
//					XWPFTable newTable = mergedDoc.createTable();
//					newTable.getCTTbl().set(table.getCTTbl().copy());
//				}
                for (IBodyElement element : doc.getBodyElements()) {
                    if (element instanceof XWPFParagraph) {
                        XWPFParagraph p = (XWPFParagraph) element;
                        XWPFParagraph newP = mergedDoc.createParagraph();
                        newP.getCTP().set(p.getCTP().copy());
                    } else if (element instanceof XWPFTable) {
                        XWPFTable table = (XWPFTable) element;
                        XWPFTable newTable = mergedDoc.createTable();
                        newTable.getCTTbl().set(table.getCTTbl().copy());
                    }
                }

				// ---- Copy headers ----
				for (XWPFHeader header : doc.getHeaderList()) {
					XWPFHeader newHeader = mergedDoc.createHeader(HeaderFooterType.DEFAULT);
					for (XWPFParagraph p : header.getParagraphs()) {
						XWPFParagraph newP = newHeader.createParagraph();
						newP.getCTP().set(p.getCTP().copy());
					}
				}

				// ---- Copy footers ----
				for (XWPFFooter footer : doc.getFooterList()) {
					XWPFFooter newFooter = mergedDoc.createFooter(HeaderFooterType.DEFAULT);
					for (XWPFParagraph p : footer.getParagraphs()) {
						XWPFParagraph newP = newFooter.createParagraph();
						newP.getCTP().set(p.getCTP().copy());
					}
				}

			}
		}
		CTBody body = mergedDoc.getDocument().getBody();
		CTSectPr sectPr = body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr();
		CTPageMar pageMar = sectPr.addNewPgMar();
		pageMar.setTop(720); // 0.5 inch
		pageMar.setBottom(720); // 0.5 inch
		pageMar.setLeft(720); // 0.5 inch
		pageMar.setRight(720); // 0.5 inch
		try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
			mergedDoc.write(baos);
			return baos.toByteArray();
		}
	}

    
}
