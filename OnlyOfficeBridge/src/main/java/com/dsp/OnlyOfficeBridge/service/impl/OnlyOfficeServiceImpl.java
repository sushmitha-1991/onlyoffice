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
import java.io.InputStream;
import java.net.URL;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

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

/**
 * 
 */
@Service
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
				downloadUrl = onlyOfficeUrl + "/only-office/download-empty-docx/" + fileName;
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
					return thirdPartyDocGen.genLobGeneration(key, docTemplates, refId);
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
						|| templateId.equals(synopsisHiLsTemplateId)) {
					return thirdPartyDocGen.genSynopsisGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(siRsTemplateId)) {
					return thirdPartyDocGen.genSIRSGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(pmbTemplateId)) {
					return thirdPartyDocGen.genPMBGeneration(key, docTemplates, refId);
				}
				if (templateId.equals(pmbResolTemplateId)) {
					return thirdPartyDocGen.genPMBResolGeneration(key, docTemplates, refId);
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

			if (status == 2 || status == 6) {
				String fileUrl = (String) payload.get("url");
				String key = (String) payload.get("key");

				// 1. Download file from OnlyOffice
				File tempFile = File.createTempFile("edited-doc", ".docx");
				try (InputStream in = new URL(fileUrl).openStream();
						FileOutputStream out = new FileOutputStream(tempFile)) {
					in.transferTo(out);
				}

				String documentName = key + ".docx";

				// 2. Upload to OpenKM
				String openKmResponse = utils.uploadPdfToOpenKM(tempFile, documentName, "/okm:root/only-office/", "231",
						"51303", "ApplicationInstance", "1");

				// 3. Parse OpenKM response
				ObjectMapper objectMapper = new ObjectMapper();
				JsonNode rootNode = objectMapper.readTree(openKmResponse);
				JsonNode dataNode = rootNode.path("data");
				String documentUuid = dataNode.path("documentUuid").asText();

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
				} else {
					metadata = OnlyOfficeMetadata.builder().applicationId(Long.parseLong(applicationId))
							.documentKey(documentUuid).dsp2Key(dsp2Key).downloadUrl(downloadUrl).onlyOfficeUrl(fileUrl)
							.createdBy(1L).updatedBy(1L).build();
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
