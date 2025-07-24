/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.utils;

import java.security.KeyFactory;
import java.security.PrivateKey;
import java.security.PublicKey;
import java.security.spec.PKCS8EncodedKeySpec;
import java.security.spec.X509EncodedKeySpec;
import java.util.Base64;
import java.util.Date;
import java.util.function.Function;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.stereotype.Component;

import io.jsonwebtoken.Claims;
import io.jsonwebtoken.Jwts;
import jakarta.annotation.PostConstruct;

@Component
public class JwtUtil {

	@Value("${jwt.tokenValidity}")
	private long tokenValidity;

	@Value("${jwt.privateKeyBase64}")
	private String privateKeyBase64;

	@Value("${jwt.publicKeyBase64}")
	private String publicKeyBase64;

	private PrivateKey privateKey;
	private PublicKey publicKey;

//	@Value("${jwt.refreshSecret}")
//	private String refreshSecret;
//
//	@Value("${jwt.refreshTokenValidity}")
//	private long refreshTokenValidity;

	@PostConstruct
	public void initKeys() throws Exception {
		this.privateKey = loadPrivateKeyFromBase64(privateKeyBase64);
		this.publicKey = loadPublicKeyFromBase64(publicKeyBase64);
	}

	public String extractUsername(String token) {
		return extractClaim(token, Claims::getSubject);
	}

	public <T> T extractClaim(String token, Function<Claims, T> claimsResolver) {
		final Claims claims = extractAllClaims(token);
		return claimsResolver.apply(claims);
	}

	private Claims extractAllClaims(String token) {
		return Jwts.parser().setSigningKey(publicKey).parseClaimsJws(token).getBody();
	}

//	public String generateToken(BaseUser baseUser, String ip) {
//
//		return Jwts.builder().setSubject(baseUser.getUsername()).claim(Constants.houseId, baseUser.getHouses().getId())
//				.claim(Constants.houseAbbr, baseUser.getHouses().getAbbr()).claim(Constants.IP, ip)
//				.setIssuedAt(new Date(System.currentTimeMillis()))
//				.setExpiration(new Date(System.currentTimeMillis() + tokenValidity))
//				.signWith(SignatureAlgorithm.RS256, privateKey).compact();
//	}

//	public String generateRefreshToken(BaseUser baseUser) {
//		Key key = Keys.hmacShaKeyFor(refreshSecret.getBytes());
//		return Jwts.builder().setSubject(baseUser.getUsername()).claim(Constants.houseId, baseUser.getHouses().getId())
//				.claim(Constants.houseAbbr, baseUser.getHouses().getAbbr())
//				.setIssuedAt(new Date(System.currentTimeMillis()))
//				.setExpiration(new Date(System.currentTimeMillis() + refreshTokenValidity))
////          .signWith(key, SignatureAlgorithm.HS256)
//				.signWith(SignatureAlgorithm.RS256, refreshSecret).compact();
//	}

	public Boolean validateToken(String token, UserDetails userDetails) {
		final String username = extractUsername(token);
		return (username.equals(userDetails.getUsername()) && !isTokenExpired(token));
	}

	private Boolean isTokenExpired(String token) {
		return extractExpiration(token).before(new Date());
	}

	private Date extractExpiration(String token) {
		return extractClaim(token, Claims::getExpiration);
	}

	private PrivateKey loadPrivateKeyFromBase64(String base64) throws Exception {
		byte[] decoded = Base64.getDecoder().decode(base64);
		PKCS8EncodedKeySpec keySpec = new PKCS8EncodedKeySpec(decoded);
		return KeyFactory.getInstance("RSA").generatePrivate(keySpec);
	}

	private PublicKey loadPublicKeyFromBase64(String base64) throws Exception {
		byte[] decoded = Base64.getDecoder().decode(base64);
		X509EncodedKeySpec keySpec = new X509EncodedKeySpec(decoded);
		return KeyFactory.getInstance("RSA").generatePublic(keySpec);
	}

}