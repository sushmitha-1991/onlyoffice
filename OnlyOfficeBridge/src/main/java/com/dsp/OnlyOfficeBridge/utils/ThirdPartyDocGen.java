package com.dsp.OnlyOfficeBridge.utils;

import java.io.*;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicInteger;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

import com.dsp.OnlyOfficeBridge.client.VerbatimClient;
import com.dsp.OnlyOfficeBridge.client.dto.FullDayInformationDTO;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
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
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.Base;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.util.List;

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
	@Value("${cd.service.url}")
	private String cdUrl;
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
					map1.put("bulletinNo", map.get("bulletinNo").toString());
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
            log.info("before merge >>>>>");
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			log.info("hhhh>>> {}",e.getMessage());
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
	
	public ResponseEntity<Resource> genCdGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = cdUrl + "/debate/lang-inputs/input-tab-view-status?status=Submitted&id=" + refId;
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

			if (docTemplate.getCode().equals("cd_merge_con")) {
				System.out.println("inside cd_merge_con");
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
	                
	                String slotTurn = data.get("slotTurn") != null ? (String) data.get("slotTurn") : "";
	                String abbr = data.get("abbr") != null ? (String) data.get("abbr") : "";
	                String counterAbbr = data.get("counterAbbr") != null ? (String) data.get("counterAbbr") : "";

	                // If counterAbbr is not empty, append to abbr
	                if (!counterAbbr.isEmpty()) {
	                    abbr = abbr + "/" + counterAbbr;
	                }

	                map.put("slotTurn", slotTurn);
	                map.put("abbr", abbr);
	                
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
	    String[] parts = refId.split("_", 4);
	    String committeeName = parts[3];
	    List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			System.out.println("Doc Template UUID >>> " + docTemplate.getDocUuid());

			boolean isHouse = "house".equalsIgnoreCase(committeeName);
			AtomicInteger docCounter = new AtomicInteger(1);
			
			// Decide which template code should be used
			String expectedCode = isHouse ? "verbatim_hour_gen_header" : "verbatim_hour_non_house_gen_header";

			if (docTemplate.getCode().equals(expectedCode)) {
				Map<String, String> map = new HashMap<>();
				map.put("meetingDate", (String) parts[1]);
				map.put("abbr", (String) parts[2]);
				map.put("pageNo", String.valueOf(docCounter.getAndIncrement()));

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
							ByteArrayResource resourceBody = utils
									.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
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
	
	public ResponseEntity<Resource> genSowGeneration(String key, List<DocTemplate> docTemplates, String refId) {
		// Build URL (refId = "1_6")
		String[] parts = refId.split("_");
		
		Long lokSabhaId = Long.parseLong(parts[0]);
		Long sessionId = Long.parseLong(parts[1]);

		String url = lobUrl + "/master/loksabhaSession/get-lok-sabha-session/" + lokSabhaId + "/" + sessionId;
		System.out.println("Calling Lok Sabha Session API URL >>> " + url);

		// Call API
		Map<String, Object> body = null;
		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, null, Map.class);
			body = response.getBody();
			System.out.println("Response from Lok Sabha API >>> " + body);
		} catch (Exception e) {
			System.err.println("Error while calling Lok Sabha API: " + e.getMessage());
			throw e;
		}

		final Map<String, Object> finalBody = body;
		List<MargeFileDto> files = new LinkedList<>();
		// Add today's date
	    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy", Locale.ENGLISH);
	    String todayDate = LocalDate.now().format(formatter);
	    System.out.println("todays date: "+todayDate);

		docTemplates.forEach(docTemplate -> {
			System.out.println("Doc Template UUID >>> " + docTemplate.getDocUuid());

			// Handle sow_header_ls template
			if ("sow_header_ls".equals(docTemplate.getCode())) {
				Map<String, Object> data = (Map<String, Object>) finalBody.get("data");
				Map<String, String> map = new HashMap<>();
				map.put("sessionNumber", (String) data.get("sessionNumber"));
				map.put("lokSabhaNumber", (String) data.get("lokSabhaNumber"));
			    map.put("todayDate", (String) todayDate);
			    
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
		});

		// Merge and return response
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

			saveOpenKm(key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sow_merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			System.err.println("Error while merging SOW docs: " + e.getMessage());
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}
//    public static byte[] mergeDocxFiles(List<MargeFileDto> margeFileDtos) throws Exception {
//        WordprocessingMLPackage targetPkg = WordprocessingMLPackage.createPackage();
//        MainDocumentPart mainPart = targetPkg.getMainDocumentPart();
//
//        for (int i = 0; i < margeFileDtos.size(); i++) {
//            MargeFileDto fileBytes = margeFileDtos.get(i);
//
//            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
//                    new org.docx4j.openpackaging.parts.PartName("/part" + i + ".docx"));
//            afiPart.setBinaryData(fileBytes.getFile());
//
//            afiPart.setContentType(
//                    new org.docx4j.openpackaging.contenttype.ContentType(
//                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
//
//            Relationship altChunkRel = mainPart.addTargetPart(afiPart);
//
//            // Insert AltChunk into main document
//            org.docx4j.wml.CTAltChunk ac =
//                    Context.getWmlObjectFactory().createCTAltChunk();
//            ac.setId(altChunkRel.getId());
//            mainPart.addObject(ac);
//        }
//
//        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
//            targetPkg.save(baos);
//            return baos.toByteArray();
//        }
//    }


    public static byte[] mergeDocxFiles(List<MargeFileDto> files) throws Exception {
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("No files to merge");
        }

        // Load first file as the base (target)
        WordprocessingMLPackage targetPkg = WordprocessingMLPackage.load(
                new ByteArrayInputStream(files.get(0).getFile())
        );
        MainDocumentPart targetMain = targetPkg.getMainDocumentPart();

        // For every additional document, copy content into target
        for (int idx = 1; idx < files.size(); idx++) {

            WordprocessingMLPackage srcPkg = WordprocessingMLPackage.load(
                    new ByteArrayInputStream(files.get(idx).getFile())
            );
            MainDocumentPart srcMain = srcPkg.getMainDocumentPart();

            // 1) Copy images from src package into target package and build oldRelId -> newRelId map
            Map<String, String> relIdMapping = new HashMap<>(); // oldRelId -> newRelId

            for (Part p : srcPkg.getParts().getParts().values()) {
                if (p instanceof BinaryPart && p.getPartName() != null &&
                        p.getPartName().getName().startsWith("/word/media/")) {

                    BinaryPart srcBin = (BinaryPart) p;

                    // read bytes from source binary part
                    ByteArrayOutputStream imgOut = new ByteArrayOutputStream();
                    srcBin.writeDataToOutputStream(imgOut);
                    byte[] imgBytes = imgOut.toByteArray();

                    // create image in target package (this also creates a rel in the target)
                    BinaryPartAbstractImage newImgPart = BinaryPartAbstractImage.createImagePart(targetPkg, imgBytes);

                    // new rel id in target for that image (usually attached to main doc or relevant part)
                    Relationship newRel = newImgPart.getSourceRelationship();
                    String newRelId = (newRel != null) ? newRel.getId() : null;

                    // for each source relationship that pointed to this binary part, map its id -> newRelId
                    // (a binary part can be referenced from main doc, header, footer, etc.)
                    List<Relationship> srcSourceRels = srcBin.getSourceRelationships();
                    if (srcSourceRels != null) {
                        for (Relationship srcRel : srcSourceRels) {
                            if (srcRel != null && srcRel.getId() != null && newRelId != null) {
                                relIdMapping.put(srcRel.getId(), newRelId);
                            }
                        }
                    }
                }
            }

            // 2) Append body content: deep-copy each object, replace old r:embed/r:id references, then add
            if (srcMain.getContent() != null && !srcMain.getContent().isEmpty()) {
                for (Object o : srcMain.getContent()) {
                    Object clone = XmlUtils.deepCopy(o);
                    String xml = XmlUtils.marshaltoString(clone, true);

                    // replace any occurrences of old rIds with new rIds
                    for (Map.Entry<String, String> e : relIdMapping.entrySet()) {
                        String oldId = e.getKey();
                        String newId = e.getValue();
                        // replace common patterns used in drawing references
                        xml = xml.replace("r:embed=\"" + oldId + "\"", "r:embed=\"" + newId + "\"");
                        xml = xml.replace("r:id=\"" + oldId + "\"", "r:id=\"" + newId + "\"");
                    }

                    Object fixed = XmlUtils.unmarshalString(xml, Context.jc);
                    targetMain.getContent().add(fixed);
                }
            }

            // 3) Merge Footnotes (null-safe)
            FootnotesPart srcFootnotes = srcMain.getFootnotesPart();
            if (srcFootnotes != null && srcFootnotes.getContents() != null) {
                FootnotesPart targetFootnotes = targetMain.getFootnotesPart();
                if (targetFootnotes == null) {
                    // create a fresh footnotes part on target and copy contents
                    FootnotesPart newFp = new FootnotesPart();
                    newFp.setContents(srcFootnotes.getContents()); // copy whole CTFootnotes
                    targetMain.addTargetPart(newFp);
                } else {
                    if (targetFootnotes.getContents() == null) {
                        targetFootnotes.setContents(new CTFootnotes());
                    }
                    if (srcFootnotes.getContents().getFootnote() != null) {
                        targetFootnotes.getContents().getFootnote()
                                .addAll(srcFootnotes.getContents().getFootnote());
                    }
                }
            }

            // 4) Merge Endnotes (null-safe)
            EndnotesPart srcEndnotes = getEndnotesPart(srcPkg);
            EndnotesPart targetEndnotes = getEndnotesPart(targetPkg);
            if (srcEndnotes != null && srcEndnotes.getContents() != null) {
                if (targetEndnotes == null) {
                    EndnotesPart newEp = new EndnotesPart();
                    newEp.setContents(srcEndnotes.getContents());
                    targetMain.addTargetPart(newEp);
                } else {
                    if (targetEndnotes.getContents() == null) {
                        targetEndnotes.setContents(new CTEndnotes());
                    }
                    if (srcEndnotes.getContents().getEndnote() != null) {
                        targetEndnotes.getContents().getEndnote().addAll(srcEndnotes.getContents().getEndnote());
                    }
                }
            }

            // 5) Merge Comments (null-safe)
            CommentsPart srcComments = srcMain.getCommentsPart();
            if (srcComments != null) {
                CommentsPart targetComments = targetMain.getCommentsPart();
                if (targetComments == null) {
                    CommentsPart newCp = new CommentsPart();
                    if (srcComments.getContents() != null) {
                        newCp.setContents(srcComments.getContents());
                    } else {
                        newCp.setContents(new Comments());
                    }
                    targetMain.addTargetPart(newCp);
                } else {
                    if (targetComments.getContents() == null) {
                        targetComments.setContents(new Comments());
                    }
                    if (srcComments.getContents() != null && srcComments.getContents().getComment() != null) {
                        targetComments.getContents().getComment().addAll(srcComments.getContents().getComment());
                    }
                }
            }

            // 6) Merge Header/Footer: best-effort — copy jaxb, remap rIds inside header/footer XML, add as new parts
//            for (Part part : srcPkg.getParts().getParts().values()) {
//                try {
//                    if (part instanceof HeaderPart) {
//                        HeaderPart srcHeader = (HeaderPart) part;
//                        if (srcHeader.getJaxbElement() != null) {
//                            Object headerClone = XmlUtils.deepCopy(srcHeader.getJaxbElement());
//                            String headerXml = XmlUtils.marshaltoString(headerClone, true);
//                            for (Map.Entry<String, String> e : relIdMapping.entrySet()) {
//                                headerXml = headerXml.replace("r:embed=\"" + e.getKey() + "\"", "r:embed=\"" + e.getValue() + "\"");
//                                headerXml = headerXml.replace("r:id=\"" + e.getKey() + "\"", "r:id=\"" + e.getValue() + "\"");
//                            }
//                            Header hdr = (Header) XmlUtils.unmarshalString(headerXml, Context.jc);
//                            HeaderPart newHeader = new HeaderPart();
//                            newHeader.setJaxbElement(hdr);
//                            targetMain.addTargetPart(newHeader);
//                        }
//                    } else if (part instanceof FooterPart) {
//                        FooterPart srcFooter = (FooterPart) part;
//                        if (srcFooter.getJaxbElement() != null) {
//                            Object footerClone = XmlUtils.deepCopy(srcFooter.getJaxbElement());
//                            String footerXml = XmlUtils.marshaltoString(footerClone, true);
//                            for (Map.Entry<String, String> e : relIdMapping.entrySet()) {
//                                footerXml = footerXml.replace("r:embed=\"" + e.getKey() + "\"", "r:embed=\"" + e.getValue() + "\"");
//                                footerXml = footerXml.replace("r:id=\"" + e.getKey() + "\"", "r:id=\"" + e.getValue() + "\"");
//                            }
//                            Footer ftr = (Footer) XmlUtils.unmarshalString(footerXml, Context.jc);
//                            FooterPart newFooter = new FooterPart();
//                            newFooter.setJaxbElement(ftr);
//                            targetMain.addTargetPart(newFooter);
//                        }
//                    }
//                } catch (Exception ex) {
//                    // keep going on header/footer errors
//                    ex.printStackTrace();
//                }
//            }

            // done with this source doc
        } // end for each source

        // Save merged doc to bytes
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            targetPkg.save(baos);
            return baos.toByteArray();
        }
    }

    // Helper to locate EndnotesPart in a package
    private static EndnotesPart getEndnotesPart(WordprocessingMLPackage pkg) {
        for (Part p : pkg.getParts().getParts().values()) {
            if (p instanceof EndnotesPart) {
                return (EndnotesPart) p;
            }
        }
        return null;
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
		FullDayInformationDTO fullDayInformationDTO = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));
		List<String> dspKeys=fullDayInformationDTO.getDsp2Keys();

		// Collect files
		List<MargeFileDto> files = new LinkedList<>();
		try {
			Map<String, String> replacements = new HashMap<>();
			replacements.put("meetingDateFullString",(String) fullDayInformationDTO.getMeetingDate());
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_ls_front_cover.docx", replacements);
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
		FullDayInformationDTO fullDayInformationDTO = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));
		List<String> dspKeys=fullDayInformationDTO.getDsp2Keys();
		// Collect files
		List<MargeFileDto> files = new LinkedList<>();

		try {
			Map<String, String> replacements = new HashMap<>();
			replacements.put("meetingDateFullString",(String) fullDayInformationDTO.getMeetingDate());
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_rs_front_cover.docx", replacements);
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
