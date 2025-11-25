package com.dsp.OnlyOfficeBridge.dto;

import lombok.Builder;
import lombok.Data;

@Builder
@Data
public class MargeFileDto {
    byte[] file;
    boolean isBreak;
    String orientation;
}
