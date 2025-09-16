package com.dsp.OnlyOfficeBridge.generator;

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
public class TranslatorGenerator extends DocGeneratorBase implements IDocGenerator {

	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Override
	String getBaseUrl(String refId) {

//		String url = translationUrl + "/inputs/input-tab-view-status?status=Submitted&id=" + refId;

		String url = baseUrl + "/inputs/input-tab-view-status?status=Submitted&id=" + refId;
		System.out.println("url>>>>>>>" + url);
		return url;
	}
	
	@Override
	public ResponseEntity<Resource> generate(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		super.url = getBaseUrl(refId);

		System.out.println("dsp2key>>>>>>>" + dsp2Key);
	
		Map<String, Object> body = getResponse();

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
						ByteArrayResource resourceBody = 
								downloadFileFromUUID(UUID.fromString(metadataOpt.get().getDocumentKey()));

						files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
					}
				});

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
