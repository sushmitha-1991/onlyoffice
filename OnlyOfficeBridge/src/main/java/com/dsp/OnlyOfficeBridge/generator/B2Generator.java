package com.dsp.OnlyOfficeBridge.generator;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;

import com.dsp.OnlyOfficeBridge.dto.MargeFileDto;
import com.dsp.OnlyOfficeBridge.entity.DocTemplate;
import com.dsp.OnlyOfficeBridge.entity.OnlyOfficeMetadata;
import com.dsp.OnlyOfficeBridge.repository.OnlyOfficeMetadataRepository;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class B2Generator extends DocGeneratorBase implements IDocGenerator {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Override
	String getBaseUrl(String refId) {

//		String url = bulletin2Url + "/get-generate-doc/" + refId;
		String url = baseUrl + "/bulletin2/get-generate-doc/" + refId;
		System.out.println("url>>>>>>>" + url);
		return url;
	}
	
	@Override
	public ResponseEntity<Resource> generate(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		System.out.println("dsp2key>>>>>>>" + dsp2Key);

		super.url = getBaseUrl(refId);
	
		Map<String, Object> body = getResponse();

		List<MargeFileDto> files = new LinkedList<>();

		docTemplates.forEach(docTemplate -> {
			System.out.println("header uuid >>>>" + docTemplate.getDocUuid());

			if (docTemplate.getCode().equals("bull2_gen_header")) {
				Map<String, String> map = new HashMap<>();
				Map<String, Object> data = (Map<String, Object>) body.get("data");
				map.put("range", (Integer) data.get("startNum") + " - " + (Integer) data.get("endNum"));
				map.put("date", (String) data.get("forDate") + "/" + (String) data.get("forDateSaka"));
				try {
					ByteArrayResource resource = replacePlaceholders(docTemplate.getDocUuid(), map);

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
						resource = replacePlaceholders(docTemplate.getDocUuid(), map1);
						resourceLine = replacePlaceholders("line.docx", mapLine);

					} catch (Exception e) {
//                        e.printStackTrace();
						throw new RuntimeException(e);
					}
					files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());

					Optional<OnlyOfficeMetadata> metadataOpt = onlyOfficeMetadataRepository
							.findByDsp2Key((String) map.get("dsp2Key"));

					ByteArrayResource resourceBody = 
							downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));
					files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					files.add(MargeFileDto.builder().file(resourceLine.getByteArray()).isBreak(true).build());
				});
			}

			else {
				ByteArrayResource resource = null;
				try {
					Map map = new HashMap();
					resource = replacePlaceholders(docTemplate.getDocUuid(), map);
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
				files.add(MargeFileDto.builder().file(resource.getByteArray()).isBreak(false).build());
			}
		});
		
		try {
			mergeFiles(files, dsp2Key);
		} catch (Exception e) {
			log.info(e.getMessage());
		}
		return null;
	}

}
