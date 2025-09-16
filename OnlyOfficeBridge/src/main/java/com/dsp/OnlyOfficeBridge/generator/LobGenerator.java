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
import com.dsp.OnlyOfficeBridge.utils.Utils;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class LobGenerator extends DocGeneratorBase implements IDocGenerator {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Override
	String getBaseUrl(String refId) {

		String url = baseUrl + "/lob/generate-doc-id/" + refId;
		System.out.println("url>>>>>>>" + url);
		return url;
	}
	
	@Override
	public ResponseEntity<Resource> generate(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		super.url = getBaseUrl(refId);
		
		Map<String, Object> body = getResponse();
		
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

						ByteArrayResource resourceBody =
								downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

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
			mergeFiles(files, dsp2Key);
//			byte[] mergedBytes = mergeDocxFiles(files);
//			ByteArrayResource mergedResource = new ByteArrayResource(mergedBytes);
//			saveOpenKm(dsp2Key, mergedResource);
//			return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
//					.header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged.docx")
//					.contentLength(mergedResource.contentLength()).body(mergedResource);
		} catch (Exception e) {

		}
		return null;
	}

}
