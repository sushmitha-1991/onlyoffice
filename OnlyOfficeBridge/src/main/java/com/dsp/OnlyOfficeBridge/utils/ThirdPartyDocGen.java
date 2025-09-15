package com.dsp.OnlyOfficeBridge.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import com.dsp.OnlyOfficeBridge.client.VerbatimClient;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestTemplate;

import com.dsp.OnlyOfficeBridge.dto.MargeFileDto;
import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.extern.slf4j.Slf4j;

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
	@Value("${translation.service.url}")
	private String translationUrl;
	@Value("${verbatim.service.url}")
	private String verbatimUrl;
	@Value("${verbatim.hourly.service.url}")
	private String verbatimHourlyUrl;
	@Value("${verbatim.staff.service.url}")
	private String verbatimStaffUrl;
	@Value("${admin.token}")
	String token;
	@Value("${download.url}")
	String downloadKmUrl;
	@Autowired
	private VerbatimClient verbatimClient;

	public ResponseEntity<Resource> genBulletin2Generation(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = bulletin2Url + "/get-generate-doc/" + refId;
		System.out.println("url>>>>>>>" + url);
		System.out.println("dsp2key>>>>>>>" + dsp2Key);

		Map<String, Object> body;
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			System.out.println("RESPOSNE OF GET-GEN-DOC-B2: " + body);
		} catch (Exception e) {
			System.err.println("Error during API call-B2: " + e.getMessage());
			throw e;
		}

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			System.out.println("header uuid >>>>" + docTemplate.getDocUuid());

			if (docTemplate.getCode().equals("bull2_gen_header")) {
				Map<String, String> map = new HashMap<>();
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				map.put("range", (Integer) data.get("startNum") + " - " + (Integer) data.get("endNum"));
				map.put("date", (String) data.get("forDate") + "/" + (String) data.get("forDateSaka"));
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);

					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else if (docTemplate.getCode().equals("bull2_gen_con")) {
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				List<Map> inputList = (List<Map>) data.get("inputList");
				inputList.forEach(map -> {
					Map<String, String> map1 = new HashMap<>();
					map1.put("myno", String.valueOf((Integer) map.get("bulletinNo")));
					map1.put("branch", (String) map.get("officeName"));
					map1.put("subject", (String) map.get("subject"));
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
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
					
					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
							.findByDsp2Key((String) map.get("dsp2Key"));
					
					
					ByteArrayResource resourceBody = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					files.add(MargeFileDto.builder().file(resourceLine.getByteArray()).isBreak(true).build());
				});

			}

			else {
				ByteArrayResource resource = null;
				try {
					Map map = new HashMap();
					resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
				files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
			}
		});
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			log.info(e.getMessage());
		}
		return null;
	}

	public ResponseEntity<Resource> genLobGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
		String url = lobUrl + "/lob/generate-doc-id/" + refId;
		System.out.println("url>>>>>>>" + url + " ==== " + token);
		Map<String, Object> body;
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			System.out.println(body);
		} catch (Exception e) {
			System.err.println("Error during API call: " + e.getMessage());
			throw e;
		}

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			System.out.println("header uuid >>>>" + docTemplate.getDocUuid());

			if (docTemplate.getCode().equals("lob_gen_header")) {
				Map<String, String> map = new HashMap<>();
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				map.put("date", (String) data.get("forDate") + "/" + (String) data.get("forDateSaka"));
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);

					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else if (docTemplate.getCode().equals("lob_gen_con")) {
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				List<Map> inputList = (List<Map>) data.get("inputList");
				inputList.forEach(map -> {
					Map<String, String> map1 = new HashMap<>();
					if (map.get("dsp2Key") != null) {
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
						files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
						
						System.out.println(">>> Find by DSP2Key: " + map.get("dsp2Key"));
						
						Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
								.findByDsp2Key((String) map.get("dsp2Key"));
						
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
						files.add(MargeFileDto.builder().file(resourceLine.getByteArray()).isBreak(true).build());
					}
				});

			}

			else {
				ByteArrayResource resource = null;
				try {
					Map map = new HashMap();
					resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
				files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());

			}
		});
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

	public ResponseEntity<Resource> genTranslationGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = translationUrl + "/inputs/input-tab-view-status?status=Submitted&id=" + refId;
		System.out.println("url>>>>>>>" + url + " ==== " + token);
		Map<String, Object> body;
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			System.out.println(body);
		} catch (Exception e) {
			System.err.println("Error during API call: " + e.getMessage());
			throw e;
		}

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			System.out.println("header uuid >>>>" + docTemplate.getDocUuid());

			if (docTemplate.getCode().equals("trans_merge_con")) {
				System.out.println("inside trans_merge_con");
				List<Map<String, Object>> dataList = (List<Map<String, Object>>) body.get("data");
				Map<String, Object> data = dataList.get(0);
				List<Map> inputList = (List<Map>) data.get("allocations");
				inputList.forEach(map -> {
					if (map.get("dsp2Key") != null) {

						Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
								.findByDsp2Key((String) map.get("dsp2Key"));
						System.out.println("searching by dsp2key in only office meta data" + map.get("dsp2Key"));
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					}
				});

			}

		});
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
	
	public ResponseEntity<Resource> genVerbatimGeneration(String key, List<DocTemplate> docTemplates, String refId) {
	    // Build URL
	    String url;
	    if (StringUtils.containsIgnoreCase(refId, "sm")) {
	        url = verbatimStaffUrl + "/merge-file/get-by-refId/" + refId;
	    } else if (StringUtils.containsIgnoreCase(refId, "turn")) {
	        url = verbatimUrl + "/get-verbatim-turn-abbr/" + refId;
	    } else {
	        url = verbatimUrl + "/get-verbatim-dsp-key/" + refId;
	    }

	    System.out.println("Calling Verbatim API URL >>> " + url);

	    if (token == null) {
	        throw new IllegalStateException("Authorization token is missing");
	    }

	    HttpHeaders headers = new HttpHeaders();
	    headers.set("Authorization", "Bearer " + token);
	    headers.setContentType(MediaType.APPLICATION_JSON);
	    HttpEntity<Void> entity = new HttpEntity<>(headers);

	    Map<String, Object> body = null;
	    try {
	        RestTemplate restTemplate = new RestTemplate();
	        ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
	        body = response.getBody();
	        System.out.println("Response from verbatim API >>> " + body);
	    } catch (Exception e) {
	        System.err.println("Error while calling verbatim API: " + e.getMessage());
	        throw e;
	    }

	    final Map<String, Object> finalBody = body;
	    List<MargeFileDto> files = new LinkedList<>();
	    docTemplates.forEach(docTemplate -> {
	        System.out.println("Doc Template UUID >>> " + docTemplate.getDocUuid());

	        // handle /get-verbatim-turn-abbr/
	        if (url.contains("/get-verbatim-turn-abbr/")) {
	            if ("verbatim_lang_gen_header".equals(docTemplate.getCode())) {
	                Map<String, Object> data = (Map<String, Object>) finalBody.get("data");
	                Map<String, String> map = new HashMap<>();
	                map.put("slotTurn", (String) data.get("slotTurn"));
	                map.put("abbr", (String) data.get("abbr"));
	                map.put("counterAbbr", (String) data.get("counterAbbr"));
	                try {
	                    ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
	                    files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
	                } catch (Exception e) {
	                    throw new RuntimeException(e);
	                }
	            }
	        }
	        else if ("verbatim_merge_con".equals(docTemplate.getCode())) {
	            List<String> dspKeys = (List<String>) finalBody.get("data");
	            dspKeys.forEach(dspKey -> {
	                Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
	                metadataOpt.ifPresent(metadata -> {
	                    ByteArrayResource resourceBody =
	                            utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
	                    files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
	                });
	            });
	        }
	    });

	    // Merge and return response
	    try {
	        byte[] mergedBytes = mergeDocxFiles(files);
	        ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

	        saveOpenKm(key, mergedResource);

	        return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
	                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=verbatim_merged.docx")
	                .contentLength(mergedResource.contentLength()).body(mergedResource);
	    } catch (Exception e) {
	        System.err.println("Error while merging verbatim docs: " + e.getMessage());
	        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
	    }
	}
	
	public ResponseEntity<Resource> genVerbatimHourGeneration(String key, List<DocTemplate> docTemplates, String refId) {
	    // Build URL 
	    String url = verbatimHourlyUrl + "/hourly-merge-dsp2key/" + refId;
	    System.out.println("Calling Verbatim Hour API URL >>> " + url);

	    if (token == null) {
	        throw new IllegalStateException("Authorization token is missing");
	    }

	    HttpHeaders headers = new HttpHeaders();
	    headers.set("Authorization", "Bearer " + token);
	    headers.setContentType(MediaType.APPLICATION_JSON);
	    HttpEntity<Void> entity = new HttpEntity<>(headers);

	    Map<String, Object> body;
	    try {
	        RestTemplate restTemplate = new RestTemplate();
	        ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
	        body = response.getBody();
	        System.out.println("Response from /get-verbatim-hour-abbr/ >>> " + body);
	    } catch (Exception e) {
	        System.err.println("Error while calling /get-verbatim-hour-abbr/: " + e.getMessage());
	        throw e;
	    }
	    String[] parts = refId.split("_", 3);
	    List<MargeFileDto> files = new LinkedList<>();

	    docTemplates.forEach(docTemplate -> {
	        System.out.println("Doc Template UUID >>> " + docTemplate.getDocUuid());
	        
	        if (docTemplate.getCode().equals("verbatim_hour_gen_header")) {
	            Map<String, String> map = new HashMap<>();
	            map.put("meetingDate", (String) parts[1]);  // placeholder
	            map.put("abbr", (String) parts[2]);         // placeholder
	            map.put("pageNo", String.valueOf(files.size() + 1));
	            try {
	                ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
	                files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
	            } catch (Exception e) {
	                throw new RuntimeException(e);
	            }

	        } else if (docTemplate.getCode().equals("verbatim_hour_gen_con")) {
	        	// data is actually a list of dspKeys
	            List<String> dspKeys = (List<String>) body.get("data");
	            if (dspKeys != null) {
	                dspKeys.forEach(dspKey -> {
	                    Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
	                    metadataOpt.ifPresent(metadata -> {
	                        ByteArrayResource resourceBody =
	                                utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
	                        files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
	                    });
	                });
	            }
	        }
	    });

	    // Merge and return response
	    try {
	        byte[] mergedBytes = mergeDocxFiles(files);
	        ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

	        saveOpenKm(key, mergedResource);

	        return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
	                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=verbatim_hour_merged.docx")
	                .contentLength(mergedResource.contentLength())
	                .body(mergedResource);
	    } catch (Exception e) {
	        System.err.println("Error while merging hour verbatim docs: " + e.getMessage());
	        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
	    }
	}




//	public ResponseEntity<Resource> genVerbatimGeneration(String key, List<DocTemplate> docTemplates, String refId) {
//		// Fetch dspKeys from your API
//		List<String> dspKeys = fetchDspKeysFromApi(refId);
//
//		// Collect files
//		List<MargeFileDto> files = new LinkedList<>();
//
//		docTemplates.forEach(docTemplate -> {
//			System.out.println("Doc Template UUID >>> " + docTemplate.getDocUuid());
//
//			if ("verbatim_merge_con".equals(docTemplate.getCode())) {
//				System.out.println("Inside verbatim_merge_con");
//
//				dspKeys.forEach(dspKey -> {
//					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
//
//					metadataOpt.ifPresent(metadata -> {
//						ByteArrayResource resourceBody = utils
//								.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
//
//						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
//					});
//				});
//			}
//		});
//
//		// Merge and return response
//		try {
//			byte[] mergedBytes = mergeDocxFiles(files);
//			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
//
//			saveOpenKm(key, mergedResource);
//
//			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
//					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=verbatim_merged.docx")
//					.contentLength(mergedResource.contentLength()).body(mergedResource);
//		} catch (Exception e) {
//			System.err.println("Error while merging verbatim docs: " + e.getMessage());
//			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
//		}
//	}
//
//	/**
//	 * Helper method to call /get-verbatim-dsp-key API and extract dspKeys.
//	 */
//	@SuppressWarnings("unchecked")
//	private List<String> fetchDspKeysFromApi(String refId) {
//		String url;
//		if (StringUtils.containsIgnoreCase(refId, "hr")) {
//			url = verbatimHourlyUrl + "/hourly-merge-dsp2key/" + refId;
//		} else if (StringUtils.containsIgnoreCase(refId, "sm")) {
//			url = verbatimStaffUrl + "/merge-file/get-by-refId/" + refId;
//		} else {
//			url = verbatimUrl + "/get-verbatim-dsp-key/" + refId;
//		}
//		System.out.println("Calling Verbatim API URL >>> " + url);
//
//		if (token == null) {
//			throw new IllegalStateException("Authorization token is missing");
//		}
//
//		HttpHeaders headers = new HttpHeaders();
//		headers.set("Authorization", "Bearer " + token);
//		headers.setContentType(MediaType.APPLICATION_JSON);
//
//		HttpEntity<Void> entity = new HttpEntity<>(headers);
//
//		try {
//			RestTemplate restTemplate = new RestTemplate();
//			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
//			System.out.println(response);
//			Map<String, Object> body = response.getBody();
//			System.out.println("Response from verbatim API >>> " + body);
//
//			if (body != null && body.get("data") != null) {
//				// ✅ data is already a List<String>
//				return (List<String>) body.get("data");
//			}
//		} catch (Exception e) {
//			System.err.println("Error while calling verbatim API: " + e.getMessage());
//			throw e;
//		}
//
//		return Collections.emptyList();
//	}

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

	public void saveOpenKm(String dsp2Key, ByteArrayResource resource) {
		try {
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
		} catch (Exception e) {
			log.info(e.getMessage());
		}

	}
	public ResponseEntity<Resource> generateVerbatimFullDayLSFile(String key, List<DocTemplate> docTemplates, String refId)  {
		List<String> dspKeys = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));

		// Collect files
		List<MargeFileDto> files = new LinkedList<>();
		try {
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_ls_front_cover.docx", new HashMap<>());
			files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(true).build());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		dspKeys.forEach(dspKey -> {
			Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);

			metadataOpt.ifPresent(metadata -> {
				ByteArrayResource resourceBody = utils
						.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));

				files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
			});
		});
		return getMergedFile(key, files);
	}


	public ResponseEntity<Resource> generateVerbatimFullDayRSFile(String key, List<DocTemplate> docTemplates, String refId) {
		List<String> dspKeys = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));

		// Collect files
		List<MargeFileDto> files = new LinkedList<>();
		try {
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_rs_front_cover.docx", new HashMap<>());
			files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(true).build());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		dspKeys.forEach(dspKey -> {
			Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);

			metadataOpt.ifPresent(metadata -> {
				ByteArrayResource resourceBody = utils
						.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));

				files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
			});
		});
		return getMergedFile(key, files);
	}

	private ResponseEntity<Resource> getMergedFile(String key, List<MargeFileDto> files) {
		// Merge and return response
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

			saveOpenKm(key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=verbatim_merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			System.err.println("Error while merging verbatim docs: " + e.getMessage());
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}
}
