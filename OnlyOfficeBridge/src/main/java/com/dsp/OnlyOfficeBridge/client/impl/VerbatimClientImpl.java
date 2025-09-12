package com.dsp.OnlyOfficeBridge.client.impl;

import com.dsp.OnlyOfficeBridge.client.VerbatimClient;
import com.dsp.OnlyOfficeBridge.response.GenericResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import java.util.List;

@Service
@Slf4j
public class VerbatimClientImpl implements VerbatimClient {

    @Value("${verbatim.ms.service.url}")
    private String verbatimServiceUrl;

    @Autowired
    private RestTemplate restTemplate;

    @Override
    public List<String> getFullFileDsp2Keys(Long meetingId) {
            String url = verbatimServiceUrl + "/proceedings/merge-of-merged/full-day-merge-dsp2key/"+meetingId;
            ParameterizedTypeReference<GenericResponse<List<String>>> responseType =
                    new ParameterizedTypeReference<>() {};

            HttpEntity<String> entity = new HttpEntity<>(null);
            ResponseEntity<GenericResponse<List<String>>> response = restTemplate.exchange(
                    url,
                    HttpMethod.GET,entity,
                    responseType
            );
            if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
                List<String> session = response.getBody().getData();
                log.info("Session Fetched: {}", session);
                return session;
            } else {
                log.error("Failed to fetch session. Response: {}", response);
                throw new RuntimeException("Could not find lastest session");
            }
    }
}
