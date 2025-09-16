package com.dsp.OnlyOfficeBridge.generator;

import org.springframework.beans.factory.ObjectProvider;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class DocGeneratorFactory { 

	@Value("${bulletin2.gen.template.id}")
	static String bulletin2TemplateId;
	
	@Value("${lob.gen.template.id}")
	static String lobTemplateId;
	
	@Value("${translation.gen.template.id}")
	static String translationTemplateId;
	
	@Value("${verbatim.gen.template.id}")
	static String verbatimTemplateId;
	
	@Value("${verbatim.lang.gen.template.id}")
	static String verbatimLangTemplateId;
	
	@Value("${verbatim.hour.gen.template.id}")
	static String verbatimHourTemplateId;
	
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
}
