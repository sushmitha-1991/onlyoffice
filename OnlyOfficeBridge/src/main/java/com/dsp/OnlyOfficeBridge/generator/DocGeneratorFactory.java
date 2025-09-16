package com.dsp.OnlyOfficeBridge.generator;

import org.springframework.beans.factory.annotation.Value;

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
		
	public static IDocGenerator getGenerator(String templateId) {
		
		if (templateId.equals(bulletin2TemplateId)) {
			return new B2Generator();
		}
		
		if (templateId.equals(lobTemplateId)) {
			return new LobGenerator();
		}
		
		if (templateId.equals(translationTemplateId)) {
			return new TranslatorGenerator();
		}
		
		if (templateId.equals(verbatimTemplateId) || templateId.equals(verbatimLangTemplateId)) {
			return new VerbatimGenerator();
		}
				
		if (templateId.equals(verbatimHourTemplateId)) {
			return new VerbatimHourlyGenerator();
		}

		return null;
	}
}
