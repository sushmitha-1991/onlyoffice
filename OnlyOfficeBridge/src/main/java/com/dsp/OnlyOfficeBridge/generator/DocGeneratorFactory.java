package com.dsp.OnlyOfficeBridge.generator;

import org.springframework.beans.factory.ObjectProvider;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import jakarta.annotation.PostConstruct;
import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class DocGeneratorFactory { 

	@Value("${bulletin2.gen.template.id}")
	String bulletin2TemplateId;
	
	@Value("${lob.gen.template.id}")
	String lobTemplateId;
	
	@Value("${translation.gen.template.id}")
	String translationTemplateId;
	
	@Value("${verbatim.gen.template.id}")
	String verbatimTemplateId;
	
	@Value("${verbatim.lang.gen.template.id}")
	String verbatimLangTemplateId;
	
	@Value("${verbatim.hour.gen.template.id}")
	String verbatimHourTemplateId;
	
	@Autowired
	private ObjectProvider<B2Generator> b2Generator;
		
	@Autowired
	private ObjectProvider<LobGenerator> lobGenerator;
	
	@Autowired
	private ObjectProvider<TranslatorGenerator> translatorGenerator;
	
	@Autowired
	private ObjectProvider<VerbatimGenerator> verbatimGenerator;
	
	@Autowired
	private ObjectProvider<VerbatimHourlyGenerator> verbatimHourlyGenerator;
	
	public IDocGenerator getGenerator(String templateId) {
		
		if (templateId.equals(bulletin2TemplateId)) {
			return b2Generator.getObject();
		}
		
		if (templateId.equals(lobTemplateId)) {
			return lobGenerator.getObject();
		}
		
		if (templateId.equals(translationTemplateId)) {
			return translatorGenerator.getObject();
		}
		
		if (templateId.equals(verbatimTemplateId) || templateId.equals(verbatimLangTemplateId)) {
			return verbatimGenerator.getObject();
		}
				
		if (templateId.equals(verbatimHourTemplateId)) {
			return verbatimHourlyGenerator.getObject();
		}

		return null;
	}
	
	   @PostConstruct
	    public void checkDependencies() {
	        log.info("Bulletin2 templateId = {}", bulletin2TemplateId);
	        log.info("LOB templateId = {}", lobTemplateId);
	        log.info("Translation templateId = {}", translationTemplateId);
	        log.info("Verbatim templateId = {}", verbatimTemplateId);
	        log.info("VerbatimLang templateId = {}", verbatimLangTemplateId);
	        log.info("VerbatimHour templateId = {}", verbatimHourTemplateId);

	        try {
	            log.info("B2Generator available? {}", b2Generator.getIfAvailable() != null);
	        } catch (Exception e) {
	            log.error("B2Generator missing!", e);
	        }

	        try {
	            log.info("LobGenerator available? {}", lobGenerator.getIfAvailable() != null);
	        } catch (Exception e) {
	            log.error("LobGenerator missing!", e);
	        }

	        try {
	            log.info("TranslatorGenerator available? {}", translatorGenerator.getIfAvailable() != null);
	        } catch (Exception e) {
	            log.error("TranslatorGenerator missing!", e);
	        }

	        try {
	            log.info("VerbatimGenerator available? {}", verbatimGenerator.getIfAvailable() != null);
	        } catch (Exception e) {
	            log.error("VerbatimGenerator missing!", e);
	        }

	        try {
	            log.info("VerbatimHourlyGenerator available? {}", verbatimHourlyGenerator.getIfAvailable() != null);
	        } catch (Exception e) {
	            log.error("VerbatimHourlyGenerator missing!", e);
	        }
	    }
	

}
