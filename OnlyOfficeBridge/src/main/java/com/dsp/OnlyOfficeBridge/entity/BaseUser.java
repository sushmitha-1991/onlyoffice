/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.entity;

import java.util.Collection;
import java.util.Collections;
import java.util.UUID;

import org.hibernate.envers.Audited;
import org.hibernate.envers.RelationTargetAuditMode;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.FetchType;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.JoinColumn;
import jakarta.persistence.ManyToOne;
import jakarta.persistence.SequenceGenerator;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = "base_user", schema = "authentication")
@Audited
public class BaseUser extends Auditable implements UserDetails {

	@Id
	@GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "base_user_id_seq")
	@SequenceGenerator(name = "base_user_id_seq", sequenceName = "authentication.base_user_id_seq", allocationSize = 1)
	private Long id;

	@Column(name = "username", unique = true, nullable = true)
	private String username;

	@Column(name = "phone_number", unique = true, nullable = false)
	private String phone;

	@Column(name = "email", unique = true, nullable = false)
	private String email;

	@Column(name = "password", nullable = false)
	private String password;

	@Column(name = "first_name", nullable = false, length = 50)
	private String firstName;

	@Column(name = "middle_name", length = 50)
	private String middleName;

	@Column(name = "abbr", length = 255)
	private String abbr;

	@Column(name = "last_name", nullable = false, length = 50)
	private String lastName;

	@Column(name = "profile_picture", columnDefinition = "uuid")
	private UUID profilePicture;

	@Column(name = "address", length = 255)
	private String address;

	@Column(name = "is_verified")
	private Boolean isVerified = false;

	@ManyToOne(fetch = FetchType.EAGER)
	@JoinColumn(name = "house_id", nullable = false, insertable = false, updatable = false)
	@Audited(targetAuditMode = RelationTargetAuditMode.NOT_AUDITED)
	private Houses house;

	@ManyToOne(fetch = FetchType.EAGER)
	@JoinColumn(name = "house_id", nullable = false, insertable = false, updatable = false)
	@Audited(targetAuditMode = RelationTargetAuditMode.NOT_AUDITED)
	private Houses houses;

	@Column(name = "house_id", nullable = false)
	private Long houseId;

	@Column(name = "designation_id", nullable = true)
	private Long designationId;

	@Column(name = "role", nullable = false)
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
		// Return true if the account is still valid (non-expired)
		return true;
	}

	@Override
	public boolean isAccountNonLocked() {
		// Return true if the account is not locked
		return true;
	}

	@Override
	public boolean isCredentialsNonExpired() {
		// Return true if the credentials are valid (non-expired)
		return true;
	}

	public Long getId() {
		return this.id;
	}

	@Override
	public boolean isEnabled() {
		// TODO Auto-generated method stub
		return false;
	}
}
