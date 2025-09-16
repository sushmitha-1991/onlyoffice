package com.dsp.OnlyOfficeBridge.generator;

import java.util.List;

import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;

import com.dsp.OnlyOfficeBridge.entity.DocTemplate;

public interface IDocGenerator {

	public ResponseEntity<Resource> generate(String dsp2Key, 
												List<DocTemplate> docTemplates, 
												String refId);
	
}
