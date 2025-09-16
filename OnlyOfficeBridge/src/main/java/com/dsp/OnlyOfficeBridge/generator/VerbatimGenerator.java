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
public class VerbatimGenerator extends DocGeneratorBase implements IDocGenerator {

	@Value("${verbatim.service.url}")
	private String verbatimUrl;
	
	@Value("${verbatim.staff.service.url}")
	private String verbatimStaffUrl;
	
	@Autowired
	private OnlyOfficeMetadataRepository onlyOfficeMetadataRepository;

	@Override
	String getBaseUrl(String refId) {

	    if (StringUtils.containsIgnoreCase(refId, "sm")) {
	        url = verbatimStaffUrl + "/merge-file/get-by-refId/" + refId;
	    } else if (StringUtils.containsIgnoreCase(refId, "turn")) {
	        url = verbatimUrl + "/get-verbatim-turn-abbr/" + refId;
	    } else {
	        url = verbatimUrl + "/get-verbatim-dsp-key/" + refId;
	    }
	    return url;
	}
	
	
	@Override
	public ResponseEntity<Resource> generate(String dsp2Key, List<DocTemplate> docTemplates, String refId) {

		System.out.println("dsp2key>>>>>>>" + dsp2Key);

		super.url = getBaseUrl(refId);
	
		Map<String, Object> body = getResponse();
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
	                            downloadFileFromUUID(UUID.fromString(metadata.getDocumentKey()));
	                    files.add(MargeFileDto.builder().file(resourceBody.getByteArray()).isBreak(false).build());
	                });
	            });
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
