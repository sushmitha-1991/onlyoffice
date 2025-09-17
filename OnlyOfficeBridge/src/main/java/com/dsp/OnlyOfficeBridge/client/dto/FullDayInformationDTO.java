package com.dsp.OnlyOfficeBridge.client.dto;

import lombok.Data;

import java.util.List;

@Data
public class FullDayInformationDTO {
    private List<String> dsp2Keys;
    private String meetingDate;
}
