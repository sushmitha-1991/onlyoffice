package com.dsp.OnlyOfficeBridge.entity;

import java.time.LocalDate;
import java.time.LocalTime;

//import com.dsp.OnlyOfficeBridge.config.XSSConfig.AllowHtml;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;


@Entity
@Table(name = "notice_details", schema = "notices")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NoticeDetails {
	
	@Id
	private Long id;
	
//	@AllowHtml
	@Column(name = "html_content", columnDefinition = "text")
	private String htmlContent;
}
