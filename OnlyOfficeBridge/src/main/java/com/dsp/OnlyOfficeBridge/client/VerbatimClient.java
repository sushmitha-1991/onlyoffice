package com.dsp.OnlyOfficeBridge.client;

import com.dsp.OnlyOfficeBridge.client.dto.FullDayInformationDTO;

import java.util.List;

public interface VerbatimClient {
    FullDayInformationDTO getFullFileDsp2Keys(Long meetingId);
}
