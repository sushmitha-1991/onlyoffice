/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.service;

import java.util.Map;

import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

/**
 * 
 */
public interface OnlyOfficeService {

	Map<String, Object> getEditorConfig(Map<String, String> payload);

	ResponseEntity<Resource> downloadByKey(String key,String templateId,String refId);
	
	public Boolean convertDocxToText(String dspKey);
	
	ResponseEntity<Resource> downloadByKeyPc(String key,String templateId,String refId,String content);
	
	ResponseEntity<Resource> downloadByKeyV2(String key,String templateId,String refId);

	ResponseEntity<Map<String, Object>> handleCallback(String dsp2Key,String applicationId, Map<String, Object> payload);
    void convertHtml();
}
