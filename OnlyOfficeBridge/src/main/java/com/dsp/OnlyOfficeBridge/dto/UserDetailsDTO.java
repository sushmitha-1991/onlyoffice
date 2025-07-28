/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.dto;

import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.stream.Collectors;

import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;

import jakarta.validation.constraints.Size;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * - * Data Transfer Object for UserDetails implementation for Spring Security -
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class UserDetailsDTO implements UserDetails {

	private Long id;
	private String username;

	@Size(min = 10, max = 10, message = "Phone number must be 10 digits")
	private String phone;

	private String email;
	private String password;
	private Long houseId;
	private boolean isActive;

	private String houseAbbr;
	private String role;

	@Override
	public Collection<? extends GrantedAuthority> getAuthorities() {
		if (this.role == null || this.role.isBlank()) {
			return Collections.emptyList();
		}

		return Arrays.stream(this.role.split(",")).map(String::trim).filter(r -> !r.isEmpty())
				.map(r -> new SimpleGrantedAuthority("ROLE_" + r.toUpperCase())).collect(Collectors.toSet());
	}

	@Override
	public boolean isAccountNonExpired() {
		// Returns true if the account is not expired
		return true;
	}

	@Override
	public boolean isAccountNonLocked() {
		// Returns true if the account is not locked
		return true;
	}

	@Override
	public boolean isCredentialsNonExpired() {
		// Returns true if the credentials are not expired
		return true;
	}

	@Override
	public boolean isEnabled() {
		// Return true if the user is enabled
		return this.isActive;
	}

	public Long geId() {
		return this.id;
	}
}
