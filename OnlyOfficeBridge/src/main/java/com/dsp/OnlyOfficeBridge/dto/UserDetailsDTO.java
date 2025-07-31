/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.dto;

import java.util.Collection;
import java.util.Collections;

import org.springframework.security.core.GrantedAuthority;
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
		// Here, you should return the roles and authorities granted to the user.
		// If you're not using GrantedAuthority, you can return an empty list or
		// customize as needed.
		return Collections.emptyList();
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
