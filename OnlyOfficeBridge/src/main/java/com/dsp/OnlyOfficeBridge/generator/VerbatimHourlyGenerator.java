package com.dsp.OnlyOfficeBridge.generator;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import org.apache.commons.lang3.StringUtils;
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
import com.dsp.OnlyOfficeBridge.utils.Utils;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class VerbatimHourlyGenerator extends DocGeneratorBase implements IDocGenerator {

	@Value("${verbatim.hourly.service.url}")
	private String verbatimHourlyUrl;
	
	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Override
	String getBaseUrl(String refId) {

	    String url = verbatimHourlyUrl + "/hourly-merge-dsp2key/" + refId;
	    return url;
	}
	
	
	@Override
	public ResponseEntity<Resource> generate(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		System.out.println("dsp2key>>>>>>>" + dsp2Key);

		super.url = getBaseUrl(refId);
	
		Map<String, Object> body = getResponse();

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
	                                downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
	                        files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
	                    });
	                });
	            }
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
