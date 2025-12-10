package com.dsp.OnlyOfficeBridge.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ConvertHtmlToDocxRequest {

    private String htmlContent;
    private Long noticeDetailsId;

}
