/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.service.impl;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import com.dsp.OnlyOfficeBridge.entity.Acronym;
import com.dsp.OnlyOfficeBridge.repository.AcronymRepository;
import com.dsp.OnlyOfficeBridge.utils.WordMLToHtml;
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

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.generator.DocGeneratorFactory;
import com.dsp.OnlyOfficeBridge.generator.IDocGenerator;
import com.dsp.OnlyOfficeBridge.repository.DocTemplateRepository;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import com.dsp.OnlyOfficeBridge.service.OnlyOfficeService;
import com.dsp.OnlyOfficeBridge.utils.ThirdPartyDocGen;
import com.dsp.OnlyOfficeBridge.utils.Utils;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import java.text.Normalizer;

import lombok.extern.slf4j.Slf4j;

/**
 * 
 */
@Service
@Slf4j
public class OnlyOfficeServiceImpl implements OnlyOfficeService {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Autowired
	private DocGeneratorFactory docGenFactory;

	@Value("${onlyoffice.url}")
	private String onlyOfficeUrl;
	@Value("${bulletin2.gen.template.id}")
	String bulletin2TemplateId;
	@Value("${bulletin2.gen.rs.template.id}")
	String bulletin2RsTemplateId;
	@Value("${lob.gen.template.id}")
	String lobTemplateId;
	@Value("${lob.gen.rs.template.id}")
	String lobRsTemplateId;
	@Value("${translation.gen.template.id}")
	String translationTemplateId;
	@Value("${verbatim.gen.template.id}")
	String verbatimTemplateId;
	@Value("${verbatim.lang.gen.template.id}")
	String verbatimLangTemplateId;
	@Value("${verbatim.hour.gen.template.id}")
	String verbatimHourTemplateId;
	@Value("${statementofwork.gen.template.id}")
	String statementOfWorkTemplateId;
	@Value("${b1.ls.gen.template.id}")
	String b1lsTemplateId;
	@Value("${b1.rs.gen.template.id}")
	String b1RsTemplateId;
	@Value("${si.rs.gen.template.id}")
	String siRsTemplateId;
	@Value("${synopsis.ls.gen.template.id}")
	String synopsisLsTemplateId;
	@Value("${synopsis.en.ls.gen.template.id}")
	String synopsisEnLsTemplateId;
	@Value("${synopsis.hi.ls.gen.template.id}")
	String synopsisHiLsTemplateId;
	@Value("${synopsis.en.rs.gen.template.id}")
	String synopsisEnRsTemplateId;
	@Value("${synopsis.hi.rs.gen.template.id}")
	String synopsisHiRsTemplateId;
	@Value("${synopsis.merge.gen.template.id}")
	String synopsisMergeTemplateId;
	@Value("${download.url}")
	String downloadKmUrl;
	@Value("${cd.gen.template.id}")
	String cdTemplateId;
	@Value("${cd.gen.merge.template.id}")
	String cdMergedTemplateId;

	@Value("${pmb.gen.template.id}")
	String pmbTemplateId;

	@Value("${pmb.resol.gen.template.id}")
	String pmbResolTemplateId;
	
	@Value("${pc.gen.template.id}")
	String pcTemplateId;
	
	@Value("${pc.merge.gen.template.id}")
	String pcMergeTemplateId;

	@Value("${sessional.resume.gen.template.id}")
	String sessionalResumeTemplateId;
	
	@Value("${resume.of.work.gen.template.id}")
	String resumeOfWorkTemplateId;
	
	@Value("${translation.version.template.id}")
	String newVersionTranslationTemplateId;
	
	@Autowired
	private Utils utils;
	@Autowired
	DocTemplateRepository docTemplateRepository;
	@Autowired
	ThirdPartyDocGen thirdPartyDocGen;
    @Autowired
    AcronymRepository acronymRepository;

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
		String emptyFlg = payload.get("empty");
		String fileType = fileName.split("\\.")[1];
		String templateId = payload.get("templateId") == null ? "0" : payload.get("templateId");
		boolean isInitStep = "INIT".equalsIgnoreCase(step);
		String keyToUse;
		String downloadUrl;
		String title;
		String dsp2Key;
		if (existingDsp2Key != null && !existingDsp2Key.isBlank()) {
			dsp2Key = existingDsp2Key;
			templateId = "0";
		} else {
			dsp2Key = UUID.randomUUID().toString();
		}

		if (isInitStep) {
			keyToUse = UUID.randomUUID().toString(); // new key
			downloadUrl = onlyOfficeUrl + "/only-office/download/1"; // static blank doc
			title = "Document.odt";
		} else {
			keyToUse = UUID.randomUUID().toString(); // new key even for edit
			if (emptyFlg.equals("1")) {
				downloadUrl = onlyOfficeUrl + "/only-office/download-empty-docx/" + fileName + "/" + dsp2Key;
			} else {
				String refId = payload.get("refId") == null ? "0" : payload.get("refId");
				downloadUrl = onlyOfficeUrl + "/only-office/download/" + dsp2Key + "/" + templateId + "/" + refId;
			}
			

			// download-empty-docx
			title = "Document-" + keyToUse + ".odt";
		}

		String callbackUrl = onlyOfficeUrl + "/only-office/callback?dsp2Key=" + dsp2Key + "&applicationId="
				+ applicationId;

		// Determine if editing is allowed
		boolean canEdit = !"false".equalsIgnoreCase(editFlag); // null and "true" are treated as editable

		// Document config
		Map<String, Object> document = new HashMap<>();
//		document.put("fileType", fileType);
		document.put("key", keyToUse);
		document.put("title", title);
		document.put("url", downloadUrl);
		document.put("permissions", Map.of("edit", canEdit));
        //document.put("options", Map.of("spellcheck", true,"lang","en-GB"));
		// Editor config
		Map<String, Object> editorConfig = new HashMap<>();
		editorConfig.put("mode", "edit");
        editorConfig.put("lang", "en-IN");
        editorConfig.put("features", Map.of("spellcheck", true));
		editorConfig.put("callbackUrl", callbackUrl);
		editorConfig.put("coEditing", Map.of("mode", "fast", "change", true));
		
//		, "zoom", 198 fit to width commented for now
		editorConfig.put("customization", Map.of("forcesave", true, "save", true,"languages",false,"spellcheck",true));
//		set default editor language to english-India
//		editorConfig.put("lang", "en-GB");
//        editorConfig.put("language","en-GB");
//        editorConfig.put("region","IN");

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
	public ResponseEntity<Resource> downloadByKey(String key, String templateId, String refId) {
		String[] keys = key.split(",");

		try {
			if (Integer.parseInt(templateId) > 0) {
				List<DocTemplate> docTemplates = docTemplateRepository
						.findByIdOrParentIdOrderBySeq(Long.valueOf(templateId), Long.parseLong(templateId));
				if (templateId.equals(bulletin2TemplateId)) {
					return thirdPartyDocGen.genBulletin2Generation(key, docTemplates, refId);
				}
				if (templateId.equals(bulletin2RsTemplateId)) {
					return thirdPartyDocGen.genBulletin2RsGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(lobTemplateId)) {
					return thirdPartyDocGen.genLobGenerationV2(key, docTemplates, refId);
				}
				if (templateId.equals(lobRsTemplateId)) {
					return thirdPartyDocGen.genLobRsGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(translationTemplateId)) {
					return thirdPartyDocGen.genTranslationGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(cdTemplateId)) {
					return thirdPartyDocGen.genCdGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(verbatimTemplateId) || templateId.equals(verbatimLangTemplateId)) {
					return thirdPartyDocGen.genVerbatimGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(verbatimFullDayLSTemplateId)) {
					return thirdPartyDocGen.generateVerbatimFullDayLSFile(key, docTemplates, refId);
				}
				if (templateId.equals(verbatimFullDayRSTemplateId)) {
					return thirdPartyDocGen.generateVerbatimFullDayRSFile(key, docTemplates, refId);
				}
				if (templateId.equals(verbatimHourTemplateId)) {
					return thirdPartyDocGen.genVerbatimHourGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(statementOfWorkTemplateId)) {
					return thirdPartyDocGen.genSowGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(cdMergedTemplateId)) {
					return thirdPartyDocGen.genCdMergedGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(b1lsTemplateId)) {
					return thirdPartyDocGen.genBulletin1LSGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(b1RsTemplateId)) {
					return thirdPartyDocGen.genBulletin1RSGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(synopsisLsTemplateId) || templateId.equals(synopsisEnLsTemplateId)
						|| templateId.equals(synopsisHiLsTemplateId) || templateId.equals(synopsisEnRsTemplateId)
						|| templateId.equals(synopsisHiRsTemplateId)) {
					return thirdPartyDocGen.genSynopsisGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(siRsTemplateId)) {
					return thirdPartyDocGen.genSIRSGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(pmbTemplateId)) {
					return thirdPartyDocGen.genPMBGenerationV2(key, docTemplates, refId);
				}
				if (templateId.equals(pmbResolTemplateId)) {
					return thirdPartyDocGen.genPMBResolGenerationV2(key, docTemplates, refId);
				}
				if (templateId.equals(sessionalResumeTemplateId) || (templateId.equals(resumeOfWorkTemplateId))) {
					return thirdPartyDocGen.genResumeOfWorkGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(newVersionTranslationTemplateId)) {
					return thirdPartyDocGen.genTranslationNewVersionGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(pcMergeTemplateId)) {
					return thirdPartyDocGen.genProvisionalCalenderFinalGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(synopsisMergeTemplateId)) {
					return thirdPartyDocGen.genMergedSynopsisGeneration(key, docTemplates, refId);
				}
			} else {
				Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
				if (metadataOpt.isPresent()) {
					ByteArrayResource resource = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
					return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
							.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
							.contentLength(resource.contentLength()).body(resource);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}

		return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
	}
	
	@Override
	public ResponseEntity<Resource> downloadByKeyPc(String key, String templateId, String refId,String content) {
		String[] keys = key.split(",");

		try {
			if (Integer.parseInt(templateId) > 0) {
				List<DocTemplate> docTemplates = docTemplateRepository
						.findByIdOrParentIdOrderBySeq(Long.valueOf(templateId), Long.parseLong(templateId));
				if (templateId.equals(pcTemplateId)) {
					return thirdPartyDocGen.genProvisionalCalenderGeneration(key, docTemplates, refId,content);
				}
				
			} else {
				Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
				if (metadataOpt.isPresent()) {
					ByteArrayResource resource = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
					return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
							.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
							.contentLength(resource.contentLength()).body(resource);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}

		return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
	}

	@Override
	public ResponseEntity<Resource> downloadByKeyV2(String key, String templateId, String refId) {

		String[] keys = key.split(",");

		try {
			if (Integer.parseInt(templateId) > 0) {

				List<DocTemplate> docTemplates = docTemplateRepository
						.findByIdOrParentIdOrderBySeq(Long.valueOf(templateId), Long.parseLong(templateId));

				if (templateId.equals(verbatimFullDayLSTemplateId)) {
					return thirdPartyDocGen.generateVerbatimFullDayLSFile(key, docTemplates, refId);
				} else if (templateId.equals(verbatimFullDayRSTemplateId)) {
					return thirdPartyDocGen.generateVerbatimFullDayRSFile(key, docTemplates, refId);
				} else {
					IDocGenerator generator = docGenFactory.getGenerator(templateId);
					return generator.generate(key, docTemplates, refId);
				}

			} else {
				Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(key);
				if (metadataOpt.isPresent()) {
					ByteArrayResource resource = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
					return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
							.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + resource.getFilename())
							.contentLength(resource.contentLength()).body(resource);
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

	// public static byte[] mergeDocxFiles(List<byte[]> docxFiles) throws Exception
	// {
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
					ByteArrayResource resource = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

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

			for (XWPFDocument doc : documents) {
				doc.getDocument().getBody().getTblArray(); // Force load
				doc.getDocument().getBody().getPArray(); // Force load
				mergedDoc.createParagraph(); // Page break between docs if needed
				mergedDoc.getDocument().getBody().addNewSectPr(); // section break
				mergedDoc.getDocument().getBody().addNewP().set(doc.getDocument().getBody().getPArray(0));
				// mergedDoc.getDocument().getBody().addNewTbl().set(
				// doc.getDocument().getBody().getTblArray(0) );
			}

			// ✅ Always set default margins (1 inch)
			utils.setDefaultMargins(mergedDoc);

			// Save merged doc to byte[]
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			mergedDoc.write(baos);

			ByteArrayResource mergedResource = new ByteArrayResource(baos.toByteArray());

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);

		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}

	private String extractUuidFromUrl(String url) {
		return url.substring(url.lastIndexOf("/") + 1);
	}

	@Override
	public ResponseEntity<Map<String, Object>> handleCallback(String dsp2Key, String applicationId,
			Map<String, Object> payload) {
		try {
			int status = (int) payload.get("status");

            System.out.println(">>> status - "+ status);

			if (status == 2 || status == 6) {
				String fileUrl = (String) payload.get("url");
				String key = (String) payload.get("key");

                System.out.println(">>> fileUrl - "+ fileUrl);
                System.out.println(">>> key - "+ key + " dspkey >>> "  + dsp2Key );
                List<String> ports = List.of("8087","8086");
                File tempFile = File.createTempFile("edited-doc", ".docx");
                for(int i=0;i<ports.size();i++)
                {
                    String url = "http://10.247.128.22:"+ports.get(i);
                    fileUrl = fileUrl.replaceAll("https://dspstg.sansad.in/doc",url);
                    try (InputStream in = new URL(fileUrl).openStream();
                         FileOutputStream out = new FileOutputStream(tempFile)) {
                        in.transferTo(out);
                        break;
                    }
                    catch (Exception e)
                    {
                        System.out.println(">>> fileUrl exception - "+ fileUrl + " what "+ e.getMessage());
                    }
                }
				// 1. Download file from OnlyOffice


				String documentName = key + ".docx";

				log.debug(">>> documentName - {}", documentName);
                System.out.println(">>> fileUrl replace - "+ fileUrl);
				// 2. Upload to OpenKM
				String openKmResponse = utils.uploadPdfToOpenKM(tempFile, documentName, "/okm:root/only-office/", "231",
						"51303", "ApplicationInstance", "1");

				log.debug(">>> openKmResponse - {}", openKmResponse);

				// 3. Parse OpenKM response
				ObjectMapper objectMapper = new ObjectMapper();
				JsonNode rootNode = objectMapper.readTree(openKmResponse);
				JsonNode dataNode = rootNode.path("data");
				String documentUuid = dataNode.path("documentUuid").asText();

				// 4. Build OpenKM download URL
				String downloadUrl = downloadKmUrl + documentUuid;

				log.debug(">>> downloadUrl - {}", downloadUrl);

				// 5. Update or insert into OnlyOfficeMetadata
				Optional<OnlyOfficeMetadata> existingMetadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dsp2Key);

				OnlyOfficeMetadata metadata;
				if (existingMetadataOpt.isPresent()) {
					log.debug(">>> inside present");
					metadata = existingMetadataOpt.get();
					metadata.setDownloadUrl(downloadUrl);
					metadata.setOnlyOfficeUrl(fileUrl);
					metadata.setUpdatedBy(1L);
					metadata.setUpdatedAt(LocalDateTime.now());
					metadata.setDocumentKey(documentUuid);
					metadata.setApplicationId(Long.parseLong(applicationId));
				} else {
					log.debug(">>> inside else");

					metadata = OnlyOfficeMetadata.builder().applicationId(Long.parseLong(applicationId))
							.documentKey(documentUuid).dsp2Key(dsp2Key).downloadUrl(downloadUrl).onlyOfficeUrl(fileUrl)
							.createdBy(1L).updatedBy(1L).build();
				}

				onlyOfficeMetadataRepository.save(metadata);

				log.debug(">>> save completed");

				// Clean up
				tempFile.delete();
			}

			return ResponseEntity.ok(Map.of("error", 0));
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(Map.of("error", 1));
		}
	}

	public Boolean convertDocxToText(String dspKey) {
        try {
            log.info("Converting DOCX to HTML for DSP Key: {}", dspKey);
 
            Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
            if (metadataOpt.isEmpty()) {
                log.warn("No metadata found for DSP Key: {}", dspKey);
                return null;
            }
            OnlyOfficeMetadata metadata = metadataOpt.get();
            ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
 
            if (resourceBody == null || resourceBody.getByteArray().length == 0) {
                log.warn("Downloaded file is empty for DSP Key: {}", dspKey);
                return null;
            }
 
            // Convert DOCX bytes to TXT
            try (ByteArrayInputStream bis = new ByteArrayInputStream(resourceBody.getByteArray());
            	     XWPFDocument document = new XWPFDocument(bis)) {
            	if(document.getParagraphs().size()>0)
            	{
            		return true;
            	}
            	else
            	{
            		return false;
            	}
//            	    StringBuilder textBuilder = new StringBuilder();

//            	    log.warn("paragraph before loop: {}", document.getParagraphs());
//            	    document.getParagraphs().forEach(paragraph -> {
//            	    	log.warn("paragraph: {}", document.getParagraphs());
//            	        String text = paragraph.getText();
//            	        log.warn("paragraph text: {}", paragraph.getText());
//            	        if (text != null && !text.trim().isEmpty()) {
//            	            textBuilder.append(text.trim()).append("\n");
//            	            log.warn("text builder append: {}", textBuilder.toString());  
//            	        }
//            	    });

//            	    String docText = textBuilder.toString().trim();

//            	    if (docText.isEmpty()) {
//            	        log.warn("Document content is empty for DSP Key: {}", dspKey);
//            	        return null; 
//            	    } else {
//            	        log.info("Document has content ({} chars)", docText.length());
//            	        log.info("Document has content :-\"{}\" ", docText);
//            	        return docText;
//            	    }

            	} catch (IOException e) {
            	    log.error("Error reading DOCX bytes for DSP Key: {}", dspKey, e);
            	    throw new RuntimeException("Failed to process DOCX file", e);
            	}
        } catch (Exception e) {
            log.error("Error converting DOCX to TEXT for DSP Key: {}", dspKey, e);
            return null;
        }
    }
	
    public void convertHtml()
    {
        List<Acronym> acronymList = acronymRepository.findAll();
        acronymList.forEach(acronym -> {
            try {
                String html = WordMLToHtml.convert(acronym.getFullText());
                acronym.setFullText(html);
                acronymRepository.save(acronym);
                log.info("save {}",acronym.getAbbr());
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

	@Override
	public byte[] convertHtmlToDocx(String htmlContent) {
		// sanitize unsafe characters
		htmlContent = sanitizeHtml(htmlContent);
		return thirdPartyDocGen.convertHtmlToDocx(htmlContent);
	}

	public static String sanitizeHtml(String html) {

		if (html == null)
			return null;

		String cleaned = html;
		// Smart quotes
		cleaned = cleaned.replace("‘", "'").replace("’", "'").replace("“", "\"").replace("”", "\"");

		// Dashes
		cleaned = cleaned.replace("–", "-").replace("—", "-");

		// Ellipsis
		cleaned = cleaned.replace("…", "...");
		return cleaned;
	}

}
