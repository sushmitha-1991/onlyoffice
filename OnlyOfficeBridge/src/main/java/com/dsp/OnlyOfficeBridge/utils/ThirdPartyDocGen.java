package com.dsp.OnlyOfficeBridge.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.BiConsumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.CommentsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EndnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTEndnotes;
import org.docx4j.wml.CTFootnotes;
import org.docx4j.wml.CTFtnEdn;
import org.docx4j.wml.Comments;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.STBrType;
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

import com.dsp.OnlyOfficeBridge.client.VerbatimClient;
import com.dsp.OnlyOfficeBridge.client.dto.FullDayInformationDTO;
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
	@Value("${cd.service.url}")
	private String cdUrl;
	@Value("${dsp2.base.url}")
	private String dsp2Url;
	@Value("${prov.calender.service.url}")
	private String pcUrl;
	@Value("${doc.convert.url}")
	private String docConvertUrl;
	@Value("${resume.of.work.service.url}")
	private String rowUrl;

	@Autowired
	private VerbatimClient verbatimClient;

	public ResponseEntity<Resource> genBulletin2Generation(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = bulletin2Url + "/get-generate-doc/" + refId;

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
		} catch (Exception e) {
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

	public ResponseEntity<Resource> genLobGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
		String url = lobUrl + "/lob/generate-doc-id/" + refId;
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
		
		// Detect House Type (LS or RS)
	    String houseType = refId != null && refId.toLowerCase().contains("_rs") ? "rs" : "ls";
	    
	    // Clean refId before sending to API
	    String cleanRefId = refId;
	    if (cleanRefId != null) {
	        cleanRefId = cleanRefId.replaceAll("(?i)_rs", "").replaceAll("(?i)_ls", "").trim();
	    }

		String url = translationUrl + "/inputs/input-tab-view-status?status=Submitted&id=" + cleanRefId;
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
		List<Map<String, Object>> dataList = (List<Map<String, Object>>) body.get("data");
		Map<String, Object> data = dataList.get(0);
		String documentType = (String) data.get("documentType");
		String documentDate = (String) data.get("documentDate");
		List<Map> inputList = (List<Map>) data.get("allocations");		
		
		// 🔹 Generate Display Date (Hindi + Saka)
	    String displayDate = getDisplayDate(documentDate, headers);
	    
		// ---------------- HEADER ADDITION ----------------
	    Map<String, String> headerInput = new HashMap<>();
	    headerInput.put("date", displayDate);

	    String headerPath = getHeaderPathByDocumentType(documentType, houseType);
	    if (headerPath != null) {
	        try {
	            ByteArrayResource headerRes = Utils.replacePlaceholders(headerPath, headerInput);
	            files.add(MargeFileDto.builder().file(headerRes.getByteArray()).isBreak(false).build());
	            System.out.println("Added header for " + documentType);
	        } catch (Exception e) {
	            System.err.println("Error adding header for " + documentType + ": " + e.getMessage());
	        }
	    }
		docTemplates.forEach(docTemplate -> {
			if (docTemplate.getCode().equals("trans_merge_con")) {
				inputList.forEach(map -> {
					if (map.get("dsp2Key") != null) {

						Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
								.findByDsp2Key((String) map.get("dsp2Key"));
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					}
				});
			}	
		});
		// ---------------- FOOTER ADDITION ----------------
	    String footerPath = getFooterPathByDocumentType(documentType, houseType);
	    String date = getDisplayDate(documentDate, headers);

	    String hindiDate = date;
	    String sakaDate = "";

	    // Check if Saka date is included
	    if (date.contains("/")) {
	        String[] parts = date.split("/");
	        hindiDate = parts[0].trim();
	        
	        if (parts.length > 1) {
	            sakaDate = parts[1].replace("(शक)", "").trim();
	        }
	    }
	    Map<String, String> footerInput = new HashMap<>();
	    footerInput.put("currentDate", hindiDate);
	    footerInput.put("currentSakaDate", sakaDate);
	    if (footerPath != null) {
	        try {
	            ByteArrayResource footerRes = Utils.replacePlaceholders(footerPath, footerInput);
	            files.add(MargeFileDto.builder().file(footerRes.getByteArray()).isBreak(false).build());
	            System.out.println("Added footer for " + documentType);
	        } catch (Exception e) {
	            System.err.println("Error adding footer for " + documentType + ": " + e.getMessage());
	        }
	    }
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
	
	// ---------------- HEADER PATH LOGIC ----------------
	private String getHeaderPathByDocumentType(String documentType, String houseType) {
		if (documentType == null)
			return null;
		String base = "translation/" + houseType + "/"; // dynamic base path
		switch (documentType.trim()) {
		case "Bulletin - Part II":
			return base + "header-b2-" + houseType + ".docx";
		case "Bulletin - Part I":
			return base + "header-b1-" + houseType + ".docx";
		case "List of Business":
			return base + "header-lob-" + houseType + ".docx";
		case "Synopsis of Debates":
			return base + "header-synopsis-" + houseType + ".docx";
		default:
			return null;
		}
	}

	// ---------------- FOOTER PATH LOGIC ----------------
	private String getFooterPathByDocumentType(String documentType, String houseType) {
		if (documentType == null)
			return null;
		String base = "translation/" + houseType + "/";
		switch (documentType.trim()) {
		case "Bulletin - Part II":
			return base + "footer-b1-b2-" + houseType + ".docx";
		case "Bulletin - Part I":
			return base + "footer-b1-b2-" + houseType + ".docx";
		case "List of Business":
			return base + "footer-lob-" + houseType + ".docx";
		case "Synopsis of Debates":
			return base + "footer-synopsis-" + houseType + ".docx";
		default:
			return null;
		}
	}
	
	private String getDisplayDate(String documentDate, HttpHeaders headers) {
	    if (documentDate == null) return "";

	    try {
	        // 🔹 Convert to LocalDate
	        LocalDate date = LocalDate.parse(documentDate);

	        // 🔹 Format in Hindi
	        DateTimeFormatter hindiFormatter = DateTimeFormatter.ofPattern("EEEE, dd MMMM, yyyy", new Locale("hi", "IN"));
	        String hindiDate = date.format(hindiFormatter);

	        // 🔹 Fetch Saka Date via API
	        String sakaDate = null;
	        try {
	            Map<String, String> requestPayload = new HashMap<>();
	            requestPayload.put("dateToConvert", documentDate);

	            RestTemplate restTemplate = new RestTemplate();
	            ResponseEntity<Map> response = restTemplate.exchange(
	            		dsp2Url + "/master/convert/to-saka-date",
	                    HttpMethod.POST,
	                    new HttpEntity<>(requestPayload, headers),
	                    Map.class
	            );

	            Map<String, Object> bodyResp = response.getBody();
	            if (bodyResp != null && bodyResp.containsKey("data")) {
	                Map<String, Object> dataResp = (Map<String, Object>) bodyResp.get("data");
	                Map<String, String> convertedDate = (Map<String, String>) dataResp.get("convertedDate");
	                sakaDate = convertedDate.get("hi");
	            }

	        } catch (Exception e) {
	            System.err.println("Error fetching Saka date: " + e.getMessage());
	        }

	        // 🔹 Combine both
	        if (sakaDate != null && !sakaDate.isEmpty()) {
	            return hindiDate + "/" + sakaDate + "(शक)";
	        } else {
	            return hindiDate;
	        }

	    } catch (Exception e) {
	        System.err.println("Error formatting display date: " + e.getMessage());
	        return documentDate; // fallback
	    }
	}



	public ResponseEntity<Resource> genCdGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
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

	public ResponseEntity<Resource> genCdMergedGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		Long historyId = onlyOfficeMetadataRepository.getHistoryId(refId);
		String url = cdUrl + "/debate/history/get/" + historyId;
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

			if (docTemplate.getCode().equals("cd_static_con")) {
				Map<String, Object> result = (Map<String, Object>) body.get("data");
				Map<String, Object> applicationInstanceDTO = (Map<String, Object>) result.get("applicationInstanceDTO");
				String lokSabhaNo = onlyOfficeMetadataRepository
						.getLokSabhaNo(Long.valueOf(result.get("lokSabhaNo").toString()));
				String lokSabhaSessionNo = onlyOfficeMetadataRepository
						.getLokSabhaSessionNo(Long.valueOf(result.get("lokSabhaSessionNo").toString()));

				String currYear = String.valueOf(LocalDate.now().getYear());

				Map<String, String> mapinput = new HashMap<>();
				Date todaysDate = new Date();
				SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
				String formattedDate = formatter.format(todaysDate);

				LocalDate todayDate = LocalDate.now();
				DateTimeFormatter formatterDay = DateTimeFormatter.ofPattern("EEEE, MMMM dd,yyyy");
				String formattedDateDay = todayDate.format(formatterDay);

				mapinput.put("headerDate", formattedDateDay + " / " + convertToSakaDate(formattedDate) + " (Saka)");
				mapinput.put("loksabhanoinwords", NumberToWordsConverter.convertLokSabhaNumberToWords(lokSabhaNo));
				mapinput.put("sessioninwords", NumberToWordsConverter.convertSessionNumberToWords(lokSabhaSessionNo));
				mapinput.put("loksabhanoinwords", NumberToWordsConverter.convertLokSabhaNumberToWords(lokSabhaNo));
				mapinput.put("sessionNo", lokSabhaSessionNo);
				mapinput.put("year", currYear);
				mapinput.put("lokSabhaNo", lokSabhaNo);
				mapinput.put("docDate",
						applicationInstanceDTO.get("documentDate") != null
								? applicationInstanceDTO.get("documentDate").toString()
								: null);

				try {
					ByteArrayResource resource = Utils.replacePlaceholders("corrected-debates/static-content.docx",
							mapinput);

					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
				System.out.println("inside cd_static_con");
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				if (data.get("dsp2Key") != null) {

					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
							.findByDsp2Key((String) data.get("dsp2Key"));
					System.out.println("searching by dsp2key in only office meta data" + data.get("dsp2Key"));
					ByteArrayResource resourceBody = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
				}

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
			System.out.println(">>> url - " + url);
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
		} catch (Exception e) {
			System.err.println("Error while calling verbatim API: " + e.getMessage());
			throw e;
		}

		final Map<String, Object> finalBody = body;
		List<MargeFileDto> files = new LinkedList<>();
		docTemplates.forEach(docTemplate -> {

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
			} else if ("verbatim_merge_con".equals(docTemplate.getCode())) {
				int[] index = {1}; // to track the current position
				List<String> dspKeys = (List<String>) finalBody.get("data");
				dspKeys.forEach(dspKey -> {
					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
					metadataOpt.ifPresent(metadata -> {
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
						if (index[0]++ == dspKeys.size()) {
							files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
						} else {
							files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
						}
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

	public ResponseEntity<Resource> genVerbatimHourGeneration(String key, List<DocTemplate> docTemplates,
			String refId) {
		// Build URL
		String url = verbatimHourlyUrl + "/hourly-merge-dsp2key/" + refId;

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
		} catch (Exception e) {
			throw e;
		}
		String[] parts = refId.split("_", 4);
		String committeeName = parts[3];
		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {

			boolean isHouse = "house".equalsIgnoreCase(committeeName);
//			AtomicInteger docCounter = new AtomicInteger(1);

			// Decide which template code should be used
			String expectedCode = isHouse ? "verbatim_hour_gen_header" : "verbatim_hour_non_house_gen_header";

			if (docTemplate.getCode().equals(expectedCode)) {
				Map<String, String> map = new HashMap<>();
				map.put("meetingDate", parts[1]);
				map.put("abbr", parts[2]);
//				map.put("pageNo", String.valueOf(docCounter.getAndIncrement()));

				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			} else if (docTemplate.getCode().equals("verbatim_hour_gen_con")) {
				int[] index = {1}; // to track the current position
				// data is actually a list of dspKeys
				List<String> dspKeys = (List<String>) body.get("data");
				if (dspKeys != null) {
					dspKeys.forEach(dspKey -> {
						Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
						metadataOpt.ifPresent(metadata -> {
							ByteArrayResource resourceBody = utils
									.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
							if (index[0]++ == dspKeys.size()) {
								files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
							} else {
								files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
							}
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
					.contentLength(mergedResource.contentLength()).body(mergedResource);
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

		// Call API
		Map<String, Object> body = null;
		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, null, Map.class);
			body = response.getBody();
		} catch (Exception e) {
			System.err.println("Error while calling Lok Sabha API: " + e.getMessage());
			throw e;
		}

		final Map<String, Object> finalBody = body;
		List<MargeFileDto> files = new LinkedList<>();
		// Add today's date
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy", Locale.ENGLISH);
		String todayDate = LocalDate.now().format(formatter);

		docTemplates.forEach(docTemplate -> {

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

	public ResponseEntity<Resource> genSynopsisGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = dsp2Url + "/synopsis/synopsis/get-synopsis-dspkeys/" + refId;
		System.out.println("url>>>>>>>" + url);

		Map<String, Object> body;
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			System.out.println("API response >>>>> " + body);
		} catch (Exception e) {
			System.err.println("Error during API call Synopsis: " + e.getMessage());
			throw e;
		}

		Map<String, Object> data = (Map<String, Object>) body.get("data");
		// default locale → English
		String locale = "en";
		if (refId != null && refId.toLowerCase().startsWith("hi_")) {
			locale = "hi";
		}
		// Pick proper formatted date
		String synopsisDate = (String) data.get("synopsisDate");
		String englishDate = (String) data.get("englishDate");
		String hindiDate = (String) data.get("hindiDate");
		String formattedDate = "en".equals(locale) ? englishDate : hindiDate;
		List<String> dspKeys = (List<String>) data.get("dspKeys");

		// 🔹 Convert Gregorian → Saka date
		String sakaDate = null;
		if (synopsisDate != null) {
			try {
				Map<String, String> requestPayload = new HashMap<>();
				requestPayload.put("dateToConvert", synopsisDate);

				RestTemplate restTemplate = new RestTemplate();
				ResponseEntity<Map> response = restTemplate.exchange(dsp2Url + "/master/convert/to-saka-date",
						HttpMethod.POST, new HttpEntity<>(requestPayload, headers), Map.class);

				Map<String, Object> bodyResp = response.getBody();
				if (bodyResp != null && bodyResp.containsKey("data")) {
					Map<String, Object> dataResp = (Map<String, Object>) bodyResp.get("data");
					Map<String, String> convertedDate = (Map<String, String>) dataResp.get("convertedDate");
					
					sakaDate = convertedDate.get(locale);
				}
			} catch (Exception e) {
				System.err.println("Error fetching Saka date: " + e.getMessage());
			}
		}

		// 🔹 Final displayDate = Gregorian + optional Saka
		final String displayDate;
		if (sakaDate != null) {
			if ("hi".equals(locale)) {
		        displayDate = formattedDate + " / " + sakaDate + " (शक)";
		    } else {
		        displayDate = formattedDate + " / " + sakaDate + " (Saka)";
		    }
		} else {
			displayDate = formattedDate;
		}
		boolean isSupplementary = refId.toLowerCase().endsWith("_sup");
		List<MargeFileDto> files = new LinkedList<>();
		docTemplates.forEach(docTemplate -> {
			// HEADER
			if (docTemplate.getCode().startsWith("synopsis_ls_gen_header") || docTemplate.getCode().startsWith("synopsis_rs_gen_header")) {

				Map<String, String> map = new HashMap<>();
				map.put("synopsisDate", (String) displayDate);

				try {
					ByteArrayResource resource = null;
					if(isSupplementary) {
						if(refId.contains("en")) {
							resource = Utils.replacePlaceholders("synopsis/en-header-synopsis-ls-sup.docx",	map);
						} else if (refId.contains("hi")) {
							resource = Utils.replacePlaceholders("synopsis/hi-header-synopsis-ls-sup.docx",	map);
						}
					} else {
						resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					}
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
			// CONTENT
			else if (docTemplate.getCode().startsWith("synopsis_ls_gen_con") || docTemplate.getCode().startsWith("synopsis_rs_gen_con")) {

				dspKeys.forEach(dspKey -> {
					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
					metadataOpt.ifPresent(metadata -> {
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					});
				});
			}
			// FOOTER
			else if (docTemplate.getCode().startsWith("synopsis_ls_gen_footer") || docTemplate.getCode().startsWith("synopsis_rs_gen_footer")) {
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
			// FALLBACK (any other template)
			else {
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
		});

		// MERGE & SAVE
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

			saveOpenKm(dsp2Key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=synopsis_merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			System.err.println("Error while merging Synopsis docs: " + e.getMessage());
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}
	
	public ResponseEntity<Resource> genMergedSynopsisGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = dsp2Url + "/synopsis/synopsis/get-merged-synopsis-dspkeys/" + refId;
		System.out.println("url>>>>>>>" + url);

		Map<String, Object> body;
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}
		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		try {
			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
			System.out.println("API response >>>>> " + body);
		} catch (Exception e) {
			System.err.println("Error during API call Synopsis: " + e.getMessage());
			throw e;
		}

		Map<String, Object> data = (Map<String, Object>) body.get("data");
		List<String> dspKeys = (List<String>) data.get("dspKeys");
		
		List<MargeFileDto> files = new LinkedList<>();
		int arr[] = { 1 };
		docTemplates.forEach(docTemplate -> {
			// CONTENT
			dspKeys.forEach(dspKey -> {
				Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
				metadataOpt.ifPresent(metadata -> {
					ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
					if (arr[0]++ == dspKeys.size()) {
						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					} else {
						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
					}
				});
			});
		});

		// MERGE & SAVE
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

			saveOpenKm(dsp2Key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=synopsis_merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			System.err.println("Error while merging Synopsis docs: " + e.getMessage());
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

		// Load first file as base
		WordprocessingMLPackage targetPkg = WordprocessingMLPackage
				.load(new ByteArrayInputStream(files.get(0).getFile()));
		MainDocumentPart targetMain = targetPkg.getMainDocumentPart();
        if(files.get(0).isBreak())
        {
            targetMain.getContent().add(createPageBreak());
        }
		// Ensure footnotes, endnotes, comments exist
		FootnotesPart targetFootnotes = targetMain.getFootnotesPart();
		if (targetFootnotes == null) {
			targetFootnotes = new FootnotesPart();
			targetFootnotes.setContents(new CTFootnotes());
			targetMain.addTargetPart(targetFootnotes);
		}
		EndnotesPart targetEndnotes = getEndnotesPart(targetPkg);
		if (targetEndnotes == null) {
			targetEndnotes = new EndnotesPart();
			targetEndnotes.setContents(new CTEndnotes());
			targetMain.addTargetPart(targetEndnotes);
		}
		CommentsPart targetComments = targetMain.getCommentsPart();
		if (targetComments == null) {
			targetComments = new CommentsPart();
			targetComments.setContents(new Comments());
			targetMain.addTargetPart(targetComments);
		}

		// Track max IDs
		int maxFootnoteId = targetFootnotes.getContents().getFootnote().stream().map(fn -> fn.getId().intValue())
				.max(Integer::compareTo).orElse(100);
		int maxEndnoteId = targetEndnotes.getContents().getEndnote().stream().map(en -> en.getId().intValue())
				.max(Integer::compareTo).orElse(100);
		int maxCommentId = targetComments.getContents().getComment().stream().map(c -> c.getId().intValue())
				.max(Integer::compareTo).orElse(100);

		// For each source file
		for (int idx = 1; idx < files.size(); idx++) {
			WordprocessingMLPackage srcPkg = WordprocessingMLPackage
					.load(new ByteArrayInputStream(files.get(idx).getFile()));
			MainDocumentPart srcMain = srcPkg.getMainDocumentPart();

			// 1) Copy images
			Map<String, String> relIdMapping = new HashMap<>();
			for (Part p : srcPkg.getParts().getParts().values()) {
				if (p instanceof BinaryPart && p.getPartName() != null
						&& p.getPartName().getName().startsWith("/word/media/")) {
					BinaryPart srcBin = (BinaryPart) p;
					ByteArrayOutputStream imgOut = new ByteArrayOutputStream();
					srcBin.writeDataToOutputStream(imgOut);
					BinaryPartAbstractImage newImgPart = BinaryPartAbstractImage.createImagePart(targetPkg,
							imgOut.toByteArray());
					Relationship newRel = newImgPart.getSourceRelationship();
					if (newRel != null) {
						for (Relationship srcRel : srcBin.getSourceRelationships()) {
							relIdMapping.put(srcRel.getId(), newRel.getId());
						}
					}
				}
			}

			// 2) Merge Footnotes
			Map<BigInteger, BigInteger> footnoteIdMap = new LinkedHashMap<>();
			FootnotesPart srcFootnotes = srcMain.getFootnotesPart();
			if (srcFootnotes != null && srcFootnotes.getContents() != null) {
				for (CTFtnEdn fn : srcFootnotes.getContents().getFootnote()) {
					BigInteger oldId = fn.getId();
					BigInteger newId = BigInteger.valueOf(++maxFootnoteId);
					fn.setId(newId);
					footnoteIdMap.put(oldId, newId);
				}
				targetFootnotes.getContents().getFootnote().addAll(srcFootnotes.getContents().getFootnote());
			}

			// 3) Merge Endnotes
			Map<BigInteger, BigInteger> endnoteIdMap = new LinkedHashMap<>();
			EndnotesPart srcEndnotes = getEndnotesPart(srcPkg);
			if (srcEndnotes != null && srcEndnotes.getContents() != null) {
				for (CTFtnEdn en : srcEndnotes.getContents().getEndnote()) {
					BigInteger oldId = en.getId();
					BigInteger newId = BigInteger.valueOf(++maxEndnoteId);
					en.setId(newId);
					endnoteIdMap.put(oldId, newId);
				}
				targetEndnotes.getContents().getEndnote().addAll(srcEndnotes.getContents().getEndnote());
			}

			// 4) Merge Comments
			Map<BigInteger, BigInteger> commentIdMap = new LinkedHashMap<>();
			CommentsPart srcComments = srcMain.getCommentsPart();
			if (srcComments != null && srcComments.getContents() != null) {
				for (Comments.Comment c : srcComments.getContents().getComment()) {
					BigInteger oldId = c.getId();
					log.info("comment old id {}", oldId);
					BigInteger newId = BigInteger.valueOf(++maxCommentId);
					c.setId(newId);
					commentIdMap.put(oldId, newId);
				}
				targetComments.getContents().getComment().addAll(srcComments.getContents().getComment());
			}
			// 5) Append body content, fixing IDs
			log.info("relIdMapping map {}", relIdMapping);
			if (srcMain.getContent() != null) {
				for (Object o : srcMain.getContent()) {
					Object clone = XmlUtils.deepCopy(o);
					String xml = XmlUtils.marshaltoString(clone, true);

					// Replace relIds

					for (Map.Entry<String, String> e : relIdMapping.entrySet()) {
						xml = xml.replace("r:embed=\"" + e.getKey() + "\"", "r:embed=\"" + e.getValue() + "\"");
						xml = xml.replace("r:id=\"" + e.getKey() + "\"", "r:id=\"" + e.getValue() + "\"");
					}
					// Replace footnote IDs

					List<Integer> currIdFootnote = getMapKeyList(xml, "w:footnoteReference[^>]*w:id=\"(\\d+)\"");
					for (Map.Entry<BigInteger, BigInteger> e : footnoteIdMap.entrySet()) {
						if (currIdFootnote.contains(e.getKey().intValue())) {
							xml = xml.replace("w:footnoteReference w:id=\"" + e.getKey() + "\"",
									"w:footnoteReference w:id=\"" + e.getValue() + "\"");
							currIdFootnote.remove(Integer.getInteger(e.getKey().toString()));
						}
					}
					// Replace endnote IDs
					for (Map.Entry<BigInteger, BigInteger> e : endnoteIdMap.entrySet()) {
						xml = xml.replace("w:endnoteReference w:id=\"" + e.getKey() + "\"",
								"w:endnoteReference w:id=\"" + e.getValue() + "\"");
					}
					// Replace comment IDs
					List<Integer> currId = getMapKeyList(xml, "w:commentReference[^>]*w:id=\"(\\d+)\"");

					for (Map.Entry<BigInteger, BigInteger> e : commentIdMap.entrySet()) {

						if (currId.contains(e.getKey().intValue())) {
							xml = xml.replace("w:commentReference w:id=\"" + e.getKey() + "\"",
									"w:commentReference w:id=\"" + e.getValue() + "\"");
							xml = xml.replace("w:commentRangeStart w:id=\"" + e.getKey() + "\"",
									"w:commentRangeStart w:id=\"" + e.getValue() + "\"");
							xml = xml.replace("w:commentRangeEnd w:id=\"" + e.getKey() + "\"",
									"w:commentRangeEnd w:id=\"" + e.getValue() + "\"");
							currId.remove(Integer.getInteger(e.getKey().toString()));
						}

					}

					Object fixed = XmlUtils.unmarshalString(xml, Context.jc);
					targetMain.getContent().add(fixed);
				}
			}

			// 6) Copy headers/footers (simple best-effort)
			List<SectionWrapper> sections = srcPkg.getDocumentModel().getSections();
			if (sections != null) {
				HeaderFooterPolicy srcPolicy = sections.get(0).getHeaderFooterPolicy();
				if (srcPolicy != null) {
					if (srcPolicy.getDefaultHeader() != null) {
						HeaderPart clonedHeader = new HeaderPart();
						clonedHeader.setContents(XmlUtils.deepCopy(srcPolicy.getDefaultHeader().getContents()));
						targetMain.addTargetPart(clonedHeader);
					}
					if (srcPolicy.getDefaultFooter() != null) {
						FooterPart clonedFooter = new FooterPart();
						clonedFooter.setContents(XmlUtils.deepCopy(srcPolicy.getDefaultFooter().getContents()));
						targetMain.addTargetPart(clonedFooter);
					}
				}
			}

			// Optional page break
			if (files.get(idx).isBreak()) {
				targetMain.getContent().add(createPageBreak());
			}
		}

		try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
			targetPkg.save(baos);
			return baos.toByteArray();
		}
	}

	private static List<Integer> getMapKeyList(String xml, String rgx) {
		List<Integer> currId = new ArrayList<>();
		try {
			Pattern pr = Pattern.compile(rgx);
			Matcher mr = pr.matcher(xml);
			while (mr.find()) {
//                if(rgx.contains("commentReference")) {
//                    log.info("commentReference id {} xml {}", mr.group(1), xml);
//                }
				currId.add(Integer.parseInt(mr.group(1)));
			}
		} catch (Exception ignored) {
		}
		return currId;
	}

	private static P createPageBreak() {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P p = factory.createP();
		R run = factory.createR();
		Br breakObj = factory.createBr();
		breakObj.setType(STBrType.PAGE); // PAGE break
		run.getContent().add(breakObj);
		p.getContent().add(run);
		return p;
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
			String documentName = dsp2Key + ".docx";

			// 2. Upload to OpenKM
			String openKmResponse = utils.uploadPdfToOpenKM(tempFile, documentName, "/okm:root/", "231", "51303",
					"ApplicationInstance", "1");

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
				metadata.setOnlyOfficeUrl(downloadUrl);
				metadata.setUpdatedBy(1L);
				metadata.setUpdatedAt(LocalDateTime.now());
				metadata.setDocumentKey(documentUuid);
			} else {
				metadata = OnlyOfficeMetadata.builder().applicationId(0L).documentKey(documentUuid).dsp2Key(dsp2Key)
						.downloadUrl(downloadUrl).onlyOfficeUrl(downloadUrl).createdBy(1L).updatedBy(1L).build();
			}

			onlyOfficeMetadataRepository.save(metadata);

			// Clean up
			tempFile.delete();
		} catch (Exception e) {
			log.info(e.getMessage());
		}

	}

	public ResponseEntity<Resource> generateVerbatimFullDayLSFile(String key, List<DocTemplate> docTemplates,
			String refId) {
		FullDayInformationDTO fullDayInformationDTO = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));
		List<String> dspKeys = fullDayInformationDTO.getDsp2Keys();

		// Collect files
		List<MargeFileDto> files = new LinkedList<>();
		try {
			Map<String, String> replacements = new HashMap<>();
			replacements.put("meetingDateFullString", (String) fullDayInformationDTO.getMeetingDate());
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_ls_front_cover.docx",
					replacements);
			files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(true).build());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		int arr[] = { 1 };
		dspKeys.forEach(dspKey -> {
			Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);

			metadataOpt.ifPresent(metadata -> {
				ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
				if (arr[0]++ == dspKeys.size()) {
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
				} else {
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
				}
			});
		});
		return getMergedFile(key, files);
	}

	public ResponseEntity<Resource> generateVerbatimFullDayRSFile(String key, List<DocTemplate> docTemplates,
			String refId) {
		FullDayInformationDTO fullDayInformationDTO = verbatimClient.getFullFileDsp2Keys(Long.valueOf(refId));
		List<String> dspKeys = fullDayInformationDTO.getDsp2Keys();
		// Collect files
		List<MargeFileDto> files = new LinkedList<>();

		try {
			Map<String, String> replacements = new HashMap<>();
			replacements.put("meetingDateFullString", (String) fullDayInformationDTO.getMeetingDate());
			ByteArrayResource resource = Utils.replacePlaceholders("verbatim/verbatim_rs_front_cover.docx",
					replacements);
			files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(true).build());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		int arr[] = { 1 };
		dspKeys.forEach(dspKey -> {
			Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);

			metadataOpt.ifPresent(metadata -> {
				ByteArrayResource resourceBody = utils.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
				if (arr[0]++ == dspKeys.size()) {
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
				} else {
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
				}
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

	/**
	 * @param
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genBulletin2RsGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		String url = bulletin2Url + "/get-generate-doc/" + refId;

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
		} catch (Exception e) {
			System.err.println("Error during API call-B2 RS: " + e.getMessage());
			throw e;
		}

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {

			if (docTemplate.getCode().equals("bull2_gen_header_rs")) {
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

			} else if (docTemplate.getCode().equals("bull2_gen_con_rs")) {
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
	 * @param
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genLobRsGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
		String url = lobUrl + "/lob/generate-doc-id/" + refId;
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
			throw e;
		}

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {

			if (docTemplate.getCode().equals("lob_gen_header_rs")) {
				Map<String, String> map = new HashMap<>();
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				map.put("date", (String) data.get("forDate"));
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);

					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else if (docTemplate.getCode().equals("lob_gen_con_rs")) {
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

						Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
								.findByDsp2Key((String) map.get("dsp2Key"));

						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
						files.add(MargeFileDto.builder().file(resourceLine.getByteArray()).isBreak(true).build());
					}
				});

			}

			else if (docTemplate.getCode().equals("lob_gen_ft_rs")) {
				Map<String, String> map = new HashMap<>();
				// Format current date like: September 09, 2025
				DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM dd, yyyy", Locale.ENGLISH);
				map.put("currentDate", LocalDate.now().format(formatter));

				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
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

	public ResponseEntity<Resource> genBulletin1LSGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {

		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {

			if (docTemplate.getCode().equals("bull1_gen_header")) {
				Map<String, String> map = new HashMap<>();

				// Example refId: b1Init-2-2025-10-17EN-T03:11PM
				String[] parts = refId.split("-");
				String gregorianDate = null;
				String timePart = null;
				String locale = "en"; // default

				// Extract session/number
				if (parts.length > 1) {
					map.put("num", parts[1]);
				}

				// Extract Gregorian date: parts[2]-parts[3]-parts[4] (e.g., 2025-10-17)
				if (parts.length > 4) {
					String dayPart = parts[4];

					// Remove locale suffix (EN/HI)
					if (dayPart.endsWith("EN") || dayPart.endsWith("HI")) {
						locale = dayPart.substring(dayPart.length() - 2).toLowerCase(); // "en" or "hi"
						dayPart = dayPart.substring(0, dayPart.length() - 2);
					}

					gregorianDate = parts[2] + "-" + parts[3] + "-" + dayPart;
				}

				// Extract time (after T)
				if (parts.length > 5 && parts[5].startsWith("T")) {
					timePart = parts[5].substring(1); // remove the 'T'
				}

				// === Convert Gregorian → Saka Date ===
				String sakaDate = null;
				if (gregorianDate != null) {
					try {
						Map<String, String> requestPayload = new HashMap<>();
						requestPayload.put("dateToConvert", gregorianDate);

						RestTemplate restTemplate = new RestTemplate();
						ResponseEntity<Map> response = restTemplate.exchange(dsp2Url + "/master/convert/to-saka-date",
								HttpMethod.POST, new HttpEntity<>(requestPayload, headers), Map.class);

						Map<String, Object> bodyResp = response.getBody();
						if (bodyResp != null && bodyResp.containsKey("data")) {
							Map<String, Object> data = (Map<String, Object>) bodyResp.get("data");
							Map<String, String> convertedDate = (Map<String, String>) data.get("convertedDate");
							sakaDate = convertedDate.get(locale);
						}
					} catch (Exception e) {
						log.error("Error fetching Saka date: {}", e.getMessage());
					}
				}

				// === Format Gregorian date ===
				DateTimeFormatter longFormatter = DateTimeFormatter.ofPattern("EEEE, MMMM dd, yyyy", Locale.ENGLISH);
				String formattedGregorianDate;
				if (gregorianDate != null) {
					formattedGregorianDate = LocalDate.parse(gregorianDate).format(longFormatter);
				} else {
					formattedGregorianDate = LocalDate.now().format(longFormatter);
				}

				// === Final date value ===
				if (sakaDate != null) {
					map.put("date", formattedGregorianDate + " / " + sakaDate + " (Saka)");
				} else {
					map.put("date", formattedGregorianDate);
				}

				// === Format time ===
				if (timePart != null) {
					// Convert "03:11PM" → "03:11 PM"
					String formattedTime = timePart.replaceAll("(?i)(AM|PM)", " $1").toUpperCase().trim();
					map.put("currentTime", formattedTime);
				} else {
					// fallback to system time
					DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("h:mm a", Locale.ENGLISH);
					map.put("currentTime", LocalTime.now().format(timeFormatter));
				}

				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else if (docTemplate.getCode().equals("bull1_gen_con")) {
				Map<String, String> emptyMap = new HashMap<>();
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), emptyMap);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			} else {
				try {
					Map<String, String> map = new HashMap<>();
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
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
			log.error("Error merging files: {}", e.getMessage());
		}

		return null;
	}

	/**
	 * @param key
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genBulletin1RSGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {

		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			if (docTemplate.getCode().equals("bull1_gen_header_rs")) {
				Map<String, String> map = new HashMap<>();

// Extract session and date from refId: B1InitRS-264-2025-09-23
				String[] parts = refId.split("-");
				int session = 0;
				String gregorianDate = null;
				String num = null;
				String timePart = null;

				if (parts.length > 2) {
					session = Integer.parseInt(parts[1]);

					// Extract number after "B1NO"
					if (parts[2].startsWith("B1NO")) {
						num = parts[2].substring(4); // "20" from "B1NO20"
					}

					gregorianDate = parts[3] + "-" + parts[4] + "-" + parts[5];

					// Extract time part (after "T")
					if (parts[6].startsWith("T")) {
						timePart = parts[6].substring(1); // e.g., "03:11PM"
					}
				}

				map.put("num", num != null ? num : String.valueOf(session));

// Convert session to words with ordinal + "SESSION"
				map.put("numTotext", NumberToWordsConverter.convertToOrdinal(session) + " SESSION");

// Format Gregorian date
				LocalDate date = LocalDate.parse(gregorianDate);
				DateTimeFormatter longFormatter = DateTimeFormatter.ofPattern("EEEE, MMMM dd, yyyy", Locale.ENGLISH);
				map.put("datePrefixDay", date.format(longFormatter)); // Tuesday, September 23, 2025

				DateTimeFormatter withoutDayFormatter = DateTimeFormatter.ofPattern("d'%s' MMMM, yyyy", Locale.ENGLISH);
				int day = date.getDayOfMonth();
				String daySuffix = getDayOfMonthSuffix(day);
				map.put("dateW/oDay", String.format("%d%s %s", day, daySuffix,
						date.format(DateTimeFormatter.ofPattern("MMMM, yyyy")))); // 23rd September, 2025

				// Format time from refId (e.g., "03:11PM" → "03:11 PM")
				if (timePart != null) {
					// Add a space before AM/PM if missing
					String formattedTime = timePart.replaceAll("(?i)(AM|PM)", " $1").toUpperCase().trim();
					map.put("currentTime", formattedTime);
				} else {
					// Fallback to current time
					DateTimeFormatter timeFormatter = DateTimeFormatter.ofPattern("h:mm a", Locale.ENGLISH);
					map.put("currentTime", LocalTime.now().format(timeFormatter));
				}

				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else if (docTemplate.getCode().equals("bull1_gen_con_rs")) {
				Map<String, String> emptyMap = new HashMap<>();
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), emptyMap);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			} else {
				try {
					Map<String, String> map = new HashMap<>();
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
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
			log.info("Error merging files >>> {}", e.getMessage());
		}
		return null;
	}

//Helper method for day suffix
	private static String getDayOfMonthSuffix(int n) {
		if (n >= 11 && n <= 13)
			return "th";
		switch (n % 10) {
		case 1:
			return "st";
		case 2:
			return "nd";
		case 3:
			return "rd";
		default:
			return "th";
		}
	}

	/**
	 * @param key
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genSIRSGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			if (docTemplate.getCode().equals("si_header_rs")) {
				Map<String, String> map = new HashMap<>();

				// Extract session, fromDate, toDate from refId: SIRS-264-2024-05-12-2024-07-12
				String[] parts = refId.split("-");
				int session = 0;
				String fromDateStr = null;
				String toDateStr = null;

				if (parts.length >= 8) {
					session = Integer.parseInt(parts[1]);

					// fromDate = 2024-05-12
					fromDateStr = parts[2] + "-" + parts[3] + "-" + parts[4];

					// toDate = 2024-07-12
					toDateStr = parts[5] + "-" + parts[6] + "-" + parts[7];
				}

				// Put session with superscript suffix (e.g., 264ᵗʰ)
				map.put("session", session + utils.getSuperscriptOrdinal(session));

				// Format fromDate and toDate
				DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

				if (fromDateStr != null) {
					LocalDate fromDate = LocalDate.parse(fromDateStr, inputFormatter);
					map.put("fromDate", fromDate.format(DateTimeFormatter.ISO_LOCAL_DATE));
				}

				if (toDateStr != null) {
					LocalDate toDate = LocalDate.parse(toDateStr, inputFormatter);
					map.put("toDate", toDate.format(DateTimeFormatter.ISO_LOCAL_DATE));
				}

				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}

			} else {
				try {
					Map<String, String> map = new HashMap<>();
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
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
			log.info("Error merging files >>> {}", e.getMessage());
		}
		return null;
	}

	/**
	 * @param key
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genPMBGeneration(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			Map<String, String> map = new HashMap<>();

			try {
				// Templates that require API data
				if (docTemplate.getCode().equals("pmb_gen_header") || docTemplate.getCode().equals("pmb_gen_ft")
						|| docTemplate.getCode().equals("pmb_gen_con_intro_ls")
						|| docTemplate.getCode().equals("pmb_gen_con_consid_ls")
						|| docTemplate.getCode().equals("pmb_gen_con_with_ls")) {

					RestTemplate restTemplate = new RestTemplate();
					ResponseEntity<Map> response = restTemplate.exchange(
							dsp2Url + "/pmb/private-member-business/" + refId, HttpMethod.GET,
							new HttpEntity<>(headers), Map.class);

					Map<String, Object> body = response.getBody();
					if (body != null && body.containsKey("data")) {
						Map<String, Object> data = (Map<String, Object>) body.get("data");

						// Header template
						if (docTemplate.getCode().equals("pmb_gen_header")) {
							String forDate = (String) data.get("forDate");
							if (forDate != null) {
								LocalDate parsedDate = LocalDate.parse(forDate);
								DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy",
										Locale.ENGLISH);
								map.put("dateWithDay", parsedDate.format(formatter));
							}
							map.put("sakaDate", (String) data.get("forDateSaka"));
						}

						// FT template
						if (docTemplate.getCode().equals("pmb_gen_ft")) {
							DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM d, yyyy", Locale.ENGLISH);
							map.put("currentDate", LocalDate.now().format(formatter));
							map.put("sakaDate", (String) data.get("forDateSaka"));
						}

						// LS templates for PMB bills
						if (docTemplate.getCode().equals("pmb_gen_con_intro_ls")
								|| docTemplate.getCode().equals("pmb_gen_con_consid_ls")
								|| docTemplate.getCode().equals("pmb_gen_con_with_ls")) {

							List<Map<String, Object>> bills = (List<Map<String, Object>>) data
									.get("privateMemberBusinessDetails");

							// helper to build content for each category
							BiConsumer<String, String> buildCategory = (categoryName, prefix) -> {
								List<Map<String, Object>> filteredBills = bills.stream()
										.filter(b -> categoryName.equals(b.get("billCategoryName")))
										.sorted(Comparator.comparingInt(b -> (Integer) b.getOrDefault("serialNo", 0)))
										.toList();

								if (!filteredBills.isEmpty()) {
									switch (prefix) {
									case "intro":
										map.put("billsForIntroduction", "(i) Bills for Introduction");
										map.put("footerForBillForIntroduction", "ALSO to introduce the bill.");
										map.put("countForBillIntro", String.valueOf(filteredBills.size()));
										break;
									case "consid":
										map.put("billsForConsid", "(ii) Bills for Consideration and Passing");
										map.put("footerForBillForConsideration",
												"ALSO to move that the Bill be passed.");
										map.put("countForBillConsid", String.valueOf(filteredBills.size()));
										break;
									case "with":
										map.put("billsForWith", "(iii) Bills for Withdrawal");
										map.put("footerForBillForWithdrawl", "ALSO to move that the Bill be passed.");
										map.put("countForBillWith", String.valueOf(filteredBills.size()));
										break;
									}

									StringBuilder sb = new StringBuilder();
									for (Map<String, Object> b : filteredBills) {
										int serialNo = (Integer) b.getOrDefault("serialNo", 0);
										if (serialNo == 0)
											serialNo = 1;

										String memberName = (String) b.getOrDefault("memberName",
												b.get("memberFullName"));
										String billName = (String) b.getOrDefault("nameOfBill", "");
										String content = (String) b.getOrDefault("content", "");

										sb.append(serialNo).append(". ").append(memberName).append("      ")
												.append(billName).append("\t").append(content)
												.append("      							").append("\n");
									}

									switch (prefix) {
									case "intro":
										map.put("contentIntro", sb.toString());
										break;
									case "consid":
										map.put("contentConsid", sb.toString());
										break;
									case "with":
										map.put("contentWith", sb.toString());
										break;
									}

								} else {
									// Clear placeholders if no bills
									switch (prefix) {
									case "intro":
										map.put("billsForIntroduction", "");
										map.put("footerForBillForIntroduction", "");
										map.put("countForBillIntro", "");
										map.put("contentIntro", "");
										break;
									case "consid":
										map.put("billsForConsid", "");
										map.put("footerForBillForConsideration", "");
										map.put("countForBillConsid", "");
										map.put("contentConsid", "");
										break;
									case "with":
										map.put("billsForWith", "");
										map.put("footerForBillForWithdrawl", "");
										map.put("countForBillWith", "");
										map.put("contentWith", "");
										break;
									}
								}
							};

							// process category based on template
							if (docTemplate.getCode().equals("pmb_gen_con_intro_ls")) {
								buildCategory.accept("PMB_BILL_INTRODUCTION", "intro");
							} else if (docTemplate.getCode().equals("pmb_gen_con_consid_ls")) {
								buildCategory.accept("PMB_BILL_CONSIDERATION", "consid");
							} else if (docTemplate.getCode().equals("pmb_gen_con_with_ls")) {
								buildCategory.accept("PMB_BILL_WITHDRAWL", "with");
							}
						}
					}

					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());

				} else {
					// Templates that do not need API data
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				}

			} catch (Exception e) {
				throw new RuntimeException("Error processing PMB data for template: " + docTemplate.getCode(), e);
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
			log.error("Error merging files >>> {}", e.getMessage());
		}

		return null;
	}

	public ResponseEntity<Resource> genPMBGenerationV2(String dsp2Key, List<DocTemplate> docTemplates, String refId) {
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			try {
				// Templates that require API data
				if (docTemplate.getCode().equals("pmb_gen_header") || docTemplate.getCode().equals("pmb_gen_ft")
						|| docTemplate.getCode().equals("pmb_gen_con_intro_ls")
						|| docTemplate.getCode().equals("pmb_gen_con_consid_ls")
						|| docTemplate.getCode().equals("pmb_gen_con_with_ls")) {

					RestTemplate restTemplate = new RestTemplate();
					ResponseEntity<Map> response = restTemplate.exchange(
							dsp2Url + "/pmb/private-member-business/" + refId, HttpMethod.GET,
							new HttpEntity<>(headers), Map.class);

					Map<String, Object> body = response.getBody();
					if (body != null && body.containsKey("data")) {
						Map<String, Object> data = (Map<String, Object>) body.get("data");

						// ===== HEADER =====
						if (docTemplate.getCode().equals("pmb_gen_header")) {
							Map<String, String> map = new HashMap<>();
							String forDate = (String) data.get("forDate");
							if (forDate != null) {
								LocalDate parsedDate = LocalDate.parse(forDate);
								DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy",
										Locale.ENGLISH);
								map.put("dateWithDay", parsedDate.format(formatter));
							}
							map.put("sakaDate", (String) data.get("forDateSaka"));

							ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
							files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
						}

						// ===== FT =====
						if (docTemplate.getCode().equals("pmb_gen_ft")) {
							Map<String, String> map = new HashMap<>();
							DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM d, yyyy", Locale.ENGLISH);
							map.put("currentDate", LocalDate.now().format(formatter));
							map.put("sakaDate", (String) data.get("forDateSaka"));

							ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
							files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
						}

						// ===== LS Templates =====
						if (docTemplate.getCode().equals("pmb_gen_con_intro_ls")
								|| docTemplate.getCode().equals("pmb_gen_con_consid_ls")
								|| docTemplate.getCode().equals("pmb_gen_con_with_ls")) {
							List<Map<String, Object>> bills = (List<Map<String, Object>>) data
									.get("privateMemberBusinessDetails");

							BiConsumer<String, String> processCategory = (categoryName, prefix) -> {
								List<Map<String, Object>> filteredBills = bills.stream()
										.filter(b -> categoryName.equals(b.get("billCategoryName")))
										.sorted(Comparator.comparingInt(b -> (Integer) b.getOrDefault("serialNo", 0)))
										.toList();

								if (!filteredBills.isEmpty()) {
									try {
										// Header
										Map<String, String> headerMap = new HashMap<>();
										String headerText = switch (prefix) {
										case "intro" -> "(i) Bills for Introduction";
										case "consid" -> "(ii) Bills for Consideration and Passing";
										case "with" -> "(iii) Bills for Withdrawal";
										default -> "";
										};
										headerMap.put("billsFor", (String) headerText);
										System.out.println("This is the headerMap for intro: " + headerMap);

										ByteArrayResource headerRes = Utils.replacePlaceholders(
												"pmb-ls/pmb_gen_con_header_common.docx", headerMap);
										files.add(MargeFileDto.builder().file(headerRes.getByteArray()).isBreak(false)
												.build());
										// Each bill content
										for (Map<String, Object> b : filteredBills) {
											Map<String, String> contentMap = new HashMap<>();
											int serialNo = (Integer) b.getOrDefault("serialNo", 0);
											++serialNo;
											String footerText = switch (prefix) {
											case "intro" -> "ALSO to introduce the bill.";
											case "consid" -> "ALSO to move that the Bill be passed.";
											case "with" -> "ALSO to move that the bill with";
											default -> "";
											};

											contentMap.put("billNumber", (String) b.getOrDefault("billNumber", ""));
											contentMap.put("serialNo", String.valueOf(serialNo));
											contentMap.put("memberName",
													(String) b.getOrDefault("memberName", b.get("memberFullName")));
											contentMap.put("billName", (String) b.getOrDefault("nameOfBill", ""));
											contentMap.put("content", (String) b.getOrDefault("content", ""));
											contentMap.put("footerForBill", footerText);

											ByteArrayResource contentRes = Utils.replacePlaceholders(
													"pmb-ls/content-pmb-billCommon-ls.docx", contentMap);
											files.add(MargeFileDto.builder().file(contentRes.getByteArray())
													.isBreak(false).build());
										}
									} catch (Exception e) {
										throw new RuntimeException(
												"Error generating PMB documents for category: " + prefix, e);
									}
								}
							};

							// Handle all three categories
							if (docTemplate.getCode().equals("pmb_gen_con_intro_ls"))
								processCategory.accept("PMB_BILL_INTRODUCTION", "intro");
							else if (docTemplate.getCode().equals("pmb_gen_con_consid_ls"))
								processCategory.accept("PMB_BILL_CONSIDERATION", "consid");
							else if (docTemplate.getCode().equals("pmb_gen_con_with_ls"))
								processCategory.accept("PMB_BILL_WITHDRAWL", "with");
						}
					}
				} else {
					// Templates that do not need API data
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				}

			} catch (Exception e) {
				throw new RuntimeException("Error processing PMB data for template: " + docTemplate.getCode(), e);
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
			log.error("Error merging files >>> {}", e.getMessage());
		}

		return null;
	}

	/**
	 * @param key
	 * @param docTemplates
	 * @param refId
	 * @return
	 */
	public ResponseEntity<Resource> genPMBResolGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		System.out.println("#### here");
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			Map<String, String> map = new HashMap<>();

			try {
				if ("pmb_resol_gen_header".equals(docTemplate.getCode())
						|| "pmb_resol_gen_ft".equals(docTemplate.getCode())) {
					RestTemplate restTemplate = new RestTemplate();
					ResponseEntity<Map> response = restTemplate.exchange(
							dsp2Url + "/pmb/private-member-business/" + refId, HttpMethod.GET,
							new HttpEntity<>(headers), Map.class);

					Map<String, Object> body = response.getBody();
					if (body != null && body.containsKey("data")) {
						Map<String, Object> data = (Map<String, Object>) body.get("data");

						// Header-specific processing
						// Header-specific processing
						if ("pmb_resol_gen_header".equals(docTemplate.getCode())) {
							// Format date
							String forDate = (String) data.get("forDate");
							if (forDate != null) {
								LocalDate parsedDate = LocalDate.parse(forDate);
								DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy",
										Locale.ENGLISH);
								map.put("dateWithDay", parsedDate.format(formatter));
							}
							map.put("sakaDate", (String) data.get("forDateSaka"));

							// Initialize placeholders
							map.put("resolutionsContent", "");

							// Build all resolutions
							List<Map<String, Object>> details = (List<Map<String, Object>>) data
									.get("privateMemberBusinessDetails");
							if (details != null && !details.isEmpty()) {
								StringBuilder sb = new StringBuilder();

								for (Map<String, Object> detail : details) {
									String memberName = (String) detail.get("memberFullName");
									if (memberName == null || memberName.isEmpty()) {
										memberName = (String) detail.get("memberName");
									}

									Integer serialNo = (Integer) detail.get("serialNo");
									if (serialNo == null || serialNo == 0) {
										serialNo = 1;
									}

									String nameOfBill = (String) detail.get("nameOfBill");
									String content = (String) detail.get("content");

									sb.append(serialNo).append("      ").append(memberName).append(" to move:-\t")
											.append(nameOfBill).append("\r\n") // newline for bill heading
											.append(content).append("\r\n\r\n")
											.append("                                 "); // double newline after
																							// content
								}

								map.put("resolutionsContent", sb.toString().trim());
							}

						}

						// FT-specific processing
						if ("pmb_resol_gen_ft".equals(docTemplate.getCode())) {
							DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM d, yyyy", Locale.ENGLISH);
							map.put("currentDate", LocalDate.now().format(formatter));
							map.put("sakaDate", (String) data.get("forDateSaka"));
						}
					}

					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} else {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				}

			} catch (Exception e) {
				throw new RuntimeException("Error processing PMB data for template: " + docTemplate.getCode(), e);
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
			log.error("Error merging files >>> {}", e.getMessage());
		}

		return null;
	}

	public ResponseEntity<Resource> genPMBResolGenerationV2(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {
		System.out.println("#### here");
		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			try {
				if ("pmb_resol_gen_header".equals(docTemplate.getCode())
						|| "pmb_resol_gen_ft".equals(docTemplate.getCode())) {

					RestTemplate restTemplate = new RestTemplate();
					ResponseEntity<Map> response = restTemplate.exchange(
							dsp2Url + "/pmb/private-member-business/" + refId, HttpMethod.GET,
							new HttpEntity<>(headers), Map.class);

					Map<String, Object> body = response.getBody();
					if (body != null && body.containsKey("data")) {
						Map<String, Object> data = (Map<String, Object>) body.get("data");

						// ================== HEADER ==================
						if ("pmb_resol_gen_header".equals(docTemplate.getCode())) {
							Map<String, String> headerMap = new HashMap<>();

							// Format date
							String forDate = (String) data.get("forDate");
							if (forDate != null) {
								LocalDate parsedDate = LocalDate.parse(forDate);
								DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE, MMMM d, yyyy",
										Locale.ENGLISH);
								headerMap.put("dateWithDay", parsedDate.format(formatter));
							}
							headerMap.put("sakaDate", (String) data.get("forDateSaka"));

							// Add header first
							try {
								ByteArrayResource headerRes = Utils
										.replacePlaceholders("pmb-ls/pmb_resol_header_common.docx", headerMap);
								files.add(MargeFileDto.builder().file(headerRes.getByteArray()).isBreak(false).build());
							} catch (Exception e) {
								throw new RuntimeException("Error creating resolution header", e);
							}

							// Get resolutions list
							List<Map<String, Object>> details = (List<Map<String, Object>>) data
									.get("privateMemberBusinessDetails");
							if (details != null && !details.isEmpty()) {
								for (Map<String, Object> detail : details) {
									Map<String, String> contentMap = new HashMap<>();

									String memberName = (String) detail.get("memberFullName");
									if (memberName == null || memberName.isEmpty()) {
										memberName = (String) detail.get("memberName");
									}

									Integer serialNo = (Integer) detail.get("serialNo");
									if (serialNo == null || serialNo == 0)
										serialNo = 1;

									String nameOfBill = (String) detail.get("nameOfBill");
									String content = (String) detail.get("content");

									contentMap.put("serialNo", String.valueOf(serialNo));
									contentMap.put("memberName", memberName);
									contentMap.put("billName", nameOfBill != null ? nameOfBill : "");
									contentMap.put("content", content != null ? content : "");

									// Each resolution content doc
									try {
										ByteArrayResource contentRes = Utils
												.replacePlaceholders("pmb-ls/content-pmb-resolCommon.docx", contentMap);
										files.add(MargeFileDto.builder().file(contentRes.getByteArray()).isBreak(false)
												.build());
									} catch (Exception e) {
										throw new RuntimeException(
												"Error creating resolution content for serial " + serialNo, e);
									}
								}
							}
						}

						// ================== FT ==================
						if ("pmb_resol_gen_ft".equals(docTemplate.getCode())) {
							Map<String, String> map = new HashMap<>();
							DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM d, yyyy", Locale.ENGLISH);
							map.put("currentDate", (String) LocalDate.now().format(formatter));
							map.put("sakaDate", (String) data.get("forDateSaka"));

							ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
							files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
						}
					}
				} else {
					// Templates that don’t need API data
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), new HashMap<>());
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				}

			} catch (Exception e) {
				throw new RuntimeException("Error processing PMB data for template: " + docTemplate.getCode(), e);
			}
		});

		// Merge and return
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);

		} catch (Exception e) {
			log.error("Error merging files >>> {}", e.getMessage());
		}

		return null;
	}

	/* ---------------- Provisional Calender gen --------------------------- */
	public ResponseEntity<Resource> genProvisionalCalenderGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId,String content) {
//		String url = pcUrl + "/provisionalcalendar/" + refId;
//		Map<String, Object> body;
//		if (token == null) {
//			throw new IllegalStateException("Authorization token is missing");
//		}
//		HttpHeaders headers = new HttpHeaders();
//		headers.set("Authorization", "Bearer " + token);
//		headers.setContentType(org.springframework.http.MediaType.APPLICATION_JSON);
//		HttpEntity<Void> entity = new HttpEntity<>(headers);
//
//		try {
//			RestTemplate restTemplate = new RestTemplate();
//			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
//			body = response.getBody();
//			System.out.println(body);
//		} catch (Exception e) {
//			System.err.println("Error during API call: " + e.getMessage());
//			throw e;
//		}

		String imageUrl="https://dspstg.sansad.in/dms/openkm/dsp1/downloads/900e055f-6420-452e-9546-46059432ee07";
		String base64Img="";
		try (InputStream in = new URL(imageUrl).openStream()) {
		    byte[] imageBytes = in.readAllBytes();
		    String base64 = Base64.getEncoder().encodeToString(imageBytes);
		     base64Img = "data:image/png;base64," + base64;
		    System.out.println(base64Img);
		} catch (IOException e) {
		    e.printStackTrace();
		}
		List<MargeFileDto> files = new LinkedList<>();
		final AtomicReference<String> htmlRef = new AtomicReference<>();
		String html="";
		docTemplates.forEach(docTemplate -> {

			if (docTemplate.getCode().equals("provisional_calender_content")) {
//				Map<String, Object> data = (Map<String, Object>) body.get("data");
//				if (data.get("content") != null) {
//					String html = (String) data.get("content");
				if(content != null) {
					htmlRef.set(content);
				}
			}
		});
		 html = htmlRef.get();
		log.info("html before replace:- {}",html);
		
		if (html == null) {
			throw new IllegalStateException("HTML content is missing for provisional calendar");
		}
		
//		html = html.replace("\r\n", " ").replace("\n", " ").replace("\r", " ");
		html = html.replace("\n", "");
		
//		String localImagePath = "images/emblem.jpg";

//		 Replacing the URL in the src attribute
		
		html = html.replace(imageUrl,base64Img);

		html = html
	            .replaceAll("[\\u00A0\\u200B\\u200C\\u200D\\uFEFF]", "")
	            .replaceAll("\\s+", " ")
	            .trim();
		html = new String(html.getBytes(StandardCharsets.UTF_8), StandardCharsets.UTF_8);
		log.info("html after replace:- {}",html);
		try {
			// byte[] mergedBytes = mergeDocxFiles(html);
			
			ByteArrayResource mergedResource = new ByteArrayResource(convertHtmlToDocx(html));
			saveOpenKm(dsp2Key, mergedResource);
			System.out.println("******** after saving openkm resource ***********");
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=pcmerged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {
			log.error("Error merging files >>> {}", e.getMessage());
		}
		return null;
	}

	public byte[] convertHtmlToDocx(String htmlContent) {
		try {
			
			System.out.println("### inside convertUrlToDocx");

			log.info("inside convertUrlToDocx htmlContent:- {}",htmlContent);
			RestTemplate restTemplate = new RestTemplate();

			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.APPLICATION_JSON);

			// Build JSON payload
			Map<String, String> payload = new HashMap<>();
			payload.put("html", htmlContent);

			HttpEntity<Map<String, String>> request = new HttpEntity<>(payload, headers);

//			// Call the /convert-to-odt endpoint
//			ResponseEntity<byte[]> response = restTemplate.exchange("http://10.247.128.22:3001/convert-to-odt",
//					HttpMethod.POST, request, byte[].class);

			System.out.println("#########" + docConvertUrl + "/convert-to-docx");
			ResponseEntity<byte[]> response = restTemplate.exchange(docConvertUrl + "/convert-to-docx", HttpMethod.POST,
					request, byte[].class);
			System.out.println(" ********converting to odt response *****************");
			System.out.println(response.getBody());
			return response.getBody();
		} catch (Exception e) {
			throw new RuntimeException("Error converting ODT to DOC via HTTP service", e);
		}
	}

	public ResponseEntity<Resource> genProvisionalCalenderFinalGeneration(String dsp2Key,
			List<DocTemplate> docTemplates, String refId) {
		String url = pcUrl + "/provisionalcalendar/" + refId;
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
			if (docTemplate.getCode().equals("pc_merge_con")) {
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				if (data.get("dsp2Key") != null) {

					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
							.findByDsp2Key((String) data.get("dsp2Key"));
					System.out.println("searching by dsp2key in only office meta data" + data.get("dsp2Key"));
					ByteArrayResource resourceBody = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
				}
				if (data.get("qbDsp2Key") != null) {

					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
							.findByDsp2Key((String) data.get("qbDsp2Key"));
					System.out.println("searching by dsp2key in only office meta data" + data.get("qbDsp2Key"));
					ByteArrayResource resourceBody = utils
							.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
				}
			}
		});
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=pcFinalMerged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {

		}
		return null;
	}

	public ResponseEntity<Resource> genTranslationNewVersionGeneration(String dsp2Key, List<DocTemplate> docTemplates,
			String refId) {

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {

			if (docTemplate.getCode().equals("new_version_dsp")) {

				Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(refId);
				ByteArrayResource resourceBody = utils
						.downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

				files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
			}
		});
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
			saveOpenKm(dsp2Key, mergedResource);
			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Translationversion.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {

		}
		return null;
	}

	public ResponseEntity<Resource> genResumeOfWorkGeneration(String key, List<DocTemplate> docTemplates,
			String refId) {
		System.out.println("Received refId: " + refId);
		String[] refParts = refId.split("-");

		if (refParts.length < 3) {
			throw new IllegalArgumentException("Invalid refId format: " + refId);
		}

		String language = refParts[0];
		String url;
		Long houseId;
		String abbr = null;
		Long applicationId = null;

		// Detect RS Flow
		boolean isRsFlow = refParts[2].contentEquals("1");

		// Build URL based on flow
		if (isRsFlow) {
			// RS Flow: EN-TMPL_RWDOS-1
			abbr = refParts[1];
			houseId = Long.parseLong(refParts[2]);
			url = lobUrl + "/master/templates/getTemplatePlaceholdersV2?abbr=" + abbr + "&language=" + language
					+ "&houseId=" + houseId;
			System.out.println("Detected RS Flow: " + url);
		} else {
			// LS Flow: EN-87-2-32-432-342-893
			abbr = refParts[1];
			houseId = Long.parseLong(refParts[2]);
			url = lobUrl + "/master/templates/getTemplatePlaceholdersV2?abbr=" + abbr + "&language=" + language
					+ "&houseId=" + houseId;
			System.out.println("Detected LS Flow: " + url);
		}

		if (token == null) {
			throw new IllegalStateException("Authorization token is missing");
		}

		HttpHeaders headers = new HttpHeaders();
		headers.set("Authorization", "Bearer " + token);
		headers.setContentType(MediaType.APPLICATION_JSON);
		HttpEntity<Void> entity = new HttpEntity<>(headers);

		// Call getTemplatePlaceholdersV2
		Map<String, Object> body;
		try {
			RestTemplate restTemplate = new RestTemplate();
			System.out.println(">>> Calling Placeholder API: " + url);
			ResponseEntity<Map> response = restTemplate.exchange(url, HttpMethod.GET, entity, Map.class);
			body = response.getBody();
		} catch (Exception e) {
			System.err.println("Error while calling Placeholder API: " + e.getMessage());
			throw new RuntimeException("Failed to fetch placeholders from API", e);
		}

		if (body == null || body.get("data") == null) {
			throw new RuntimeException("No placeholder data received from API for " + refId);
		}

		List<?> dataList = (List<?>) body.get("data");
		if (dataList.isEmpty()) {
			throw new RuntimeException("Empty placeholder list returned from API for " + refId);
		}

		LocalDate today = LocalDate.now();

		String monthYear = today.format(DateTimeFormatter.ofPattern("MMMM, yyyy"));
		String isoDate = today.toString(); // already in yyyy-MM-dd format

		// 🔹 Convert Gregorian → Saka date
		String sakaDate = null;
		if (isoDate != null) {
			try {
				Map<String, String> requestPayload = new HashMap<>();
				requestPayload.put("dateToConvert", isoDate);

				RestTemplate restTemplate = new RestTemplate();
				ResponseEntity<Map> response = restTemplate.exchange(dsp2Url + "/master/convert/to-saka-date",
						HttpMethod.POST, new HttpEntity<>(requestPayload, headers), Map.class);

				Map<String, Object> bodyResp = response.getBody();
				if (bodyResp != null && bodyResp.containsKey("data")) {
					Map<String, Object> dataResp = (Map<String, Object>) bodyResp.get("data");
					Map<String, String> convertedDate = (Map<String, String>) dataResp.get("convertedDate");

					// default locale → English
					String locale = "en";
					if (refId.endsWith("HI")) {
						locale = "hi";
					}
					sakaDate = convertedDate.get(locale);
				}
			} catch (Exception e) {
				System.err.println("Error fetching Saka date: " + e.getMessage());
			}
		}

		// 🔹 Final displayDate = Gregorian + optional Saka
		final String displayDate;
		if (sakaDate != null) {
			displayDate = monthYear + " / " + sakaDate + " (Saka)";
		} else {
			displayDate = monthYear;
		}

		Map<String, Object> firstEntry = (Map<String, Object>) dataList.get(0);
		Map<String, Object> placeholders = (Map<String, Object>) firstEntry.get("placeholders");

		// Extract placeholders
		String sessionNumber = placeholders.getOrDefault("sessionNumber", "").toString();
		String sessionDate = placeholders.getOrDefault("sessionDate", "").toString();

		Map<String, String> map = new HashMap<>();
		map.put("sessionNumber", sessionNumber);
		map.put("sessionDate", sessionDate);
		map.put("displayDate", displayDate);
		map.put("todaysDate", monthYear);
		map.put("sakaDate", sakaDate != null ? sakaDate : "");

		// Initialize file list
		List<MargeFileDto> files = new LinkedList<>();

		// Determine correct header template based on flow
		String headerTemplateCode = isRsFlow ? "sessional_resume_gen_header" : "resume_of_work_gen_header";

		// Replace placeholders in the header template
		docTemplates.forEach(docTemplate -> {
			if (headerTemplateCode.equals(docTemplate.getCode())) {
				try {
					ByteArrayResource resource = Utils.replacePlaceholders(docTemplate.getDocUuid(), map);
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
				} catch (Exception e) {
					throw new RuntimeException("Error replacing placeholders: " + e.getMessage(), e);
				}
			}
		});

		// LS FLOW: Fetch DSP2 keys and merge documents
		if (!isRsFlow) {
			List<Long> inputIds = new ArrayList<>();
			for (int i = 3; i < refParts.length; i++) {
				try {
					inputIds.add(Long.parseLong(refParts[i]));
				} catch (NumberFormatException ignored) {
				}
			}

			System.out.println("Input IDs for DSP2 call: " + inputIds);

			if (!inputIds.isEmpty()) {
				String dsp2Url = rowUrl + "/get-dsp2keys";
				System.out.print("calling resume of work :" + dsp2Url);

				HttpEntity<List<Long>> dsp2Entity = new HttpEntity<>(inputIds, headers);
				// Call the DSP2 API
				Map<String, Object> dsp2Body;
				try {
					RestTemplate restTemplate = new RestTemplate();
					ResponseEntity<Map> dsp2Response = restTemplate.exchange(dsp2Url, HttpMethod.POST, dsp2Entity,
							Map.class);
					dsp2Body = dsp2Response.getBody();
				} catch (Exception e) {
					throw new RuntimeException("Failed to fetch DSP2 keys: " + e.getMessage(), e);
				}

				// Extract "data" directly as List<String>
				List<String> dsp2Keys = new ArrayList<>();
				if (dsp2Body != null && dsp2Body.get("data") instanceof List) {
					List<?> rawList = (List<?>) dsp2Body.get("data");
					rawList.forEach(item -> dsp2Keys.add(item.toString()));
				}

				System.out.println("DSP2 Keys fetched: " + dsp2Keys);

				// Fetch documents for each DSP2 key and add to files
				for (String dspKey : dsp2Keys) {
					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository.findByDsp2Key(dspKey);
					metadataOpt.ifPresent(metadata -> {
						ByteArrayResource resourceBody = utils
								.downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(true).build());
					});
				}
			}
		}

		// Merge all files and return DOCX
		try {
			byte[] mergedBytes = mergeDocxFiles(files);
			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);

			saveOpenKm(key, mergedResource);

			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sessional_resume_merged.docx")
					.contentLength(mergedResource.contentLength()).body(mergedResource);

		} catch (Exception e) {
			System.err.println("Error while merging sessional resume docs: " + e.getMessage());
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
		}
	}

//	Convert Gregorian → Saka date common function
	public String convertToSakaDate(String inputDate) {
		String sakaDate = null;
		if (inputDate != null) {
			try {
				HttpHeaders headers = new HttpHeaders();
				headers.set("Authorization", "Bearer " + token);
				headers.setContentType(MediaType.APPLICATION_JSON);
				HttpEntity<Void> entity = new HttpEntity<>(headers);

				Map<String, String> requestPayload = new HashMap<>();
				requestPayload.put("dateToConvert", inputDate);

				RestTemplate restTemplate = new RestTemplate();
				ResponseEntity<Map> response = restTemplate.exchange(dsp2Url + "/master/convert/to-saka-date",
						HttpMethod.POST, new HttpEntity<>(requestPayload, headers), Map.class);

				Map<String, Object> bodyResp = response.getBody();
				if (bodyResp != null && bodyResp.containsKey("data")) {
					Map<String, Object> dataResp = (Map<String, Object>) bodyResp.get("data");
					Map<String, String> convertedDate = (Map<String, String>) dataResp.get("convertedDate");

					sakaDate = convertedDate.get("en");
				}
			} catch (Exception e) {
				System.err.println("Error fetching Saka date: " + e.getMessage());
			}
		}
		return sakaDate;
	}
}
