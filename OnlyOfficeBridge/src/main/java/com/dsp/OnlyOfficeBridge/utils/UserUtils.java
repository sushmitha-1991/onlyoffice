package com.dsp.OnlyOfficeBridge.utils;

import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Component;

import com.dsp.OnlyOfficeBridge.dto.UserDetailsDTO;

@Component
public class UserUtils {

	public Long getHouseId() {
		Long houseId;
		Object principal = SecurityContextHolder.getContext().getAuthentication().getPrincipal();
		if (principal instanceof UserDetailsDTO userDetailsDTO) {
			houseId = userDetailsDTO.getHouseId();
		} else {
			throw new RuntimeException("Bad Authentication Could not find house id");
		}
		return houseId;
	}
}