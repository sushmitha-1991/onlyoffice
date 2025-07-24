/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.dsp.OnlyOfficeBridge.entity.BaseUser;

/**
 * 
 */
@Repository
public interface MyUserRepository extends JpaRepository<BaseUser, Long> {
	BaseUser findByUsername(String username);

	BaseUser findByEmail(String email);

	BaseUser findByPhone(String phone);
}