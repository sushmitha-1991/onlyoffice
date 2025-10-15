/**
 * 
 * @author Mrinal Lakhanpal
 * @version 1.0
 */
package com.dsp.OnlyOfficeBridge.utils;

import java.util.HashMap;
import java.util.Map;

import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.RestTemplate;

public class NumberToWordsConverter {

	private static final String[] tensNames = { "", " TEN", " TWENTY", " THIRTY", " FORTY", " FIFTY", " SIXTY",
			" SEVENTY", " EIGHTY", " NINETY" };

	private static final String[] numNames = { "", " ONE", " TWO", " THREE", " FOUR", " FIVE", " SIX", " SEVEN",
			" EIGHT", " NINE", " TEN", " ELEVEN", " TWELVE", " THIRTEEN", " FOURTEEN", " FIFTEEN", " SIXTEEN",
			" SEVENTEEN", " EIGHTEEN", " NINETEEN" };

	private static String convertLessThanOneThousand(int number) {
		String soFar;

		if (number % 100 < 20) {
			soFar = numNames[number % 100];
			number /= 100;
		} else {
			soFar = numNames[number % 10];
			number /= 10;

			soFar = tensNames[number % 10] + soFar;
			number /= 10;
		}
		if (number == 0)
			return soFar;
		return numNames[number] + " HUNDRED" + soFar;
	}

	public static String convert(int number) {
		if (number == 0) {
			return "ZERO";
		}

		String sNumber = String.format("%03d", number);
		int n = Integer.parseInt(sNumber);

		String result = convertLessThanOneThousand(n).trim();
		return result;
	}

	// Add ordinal suffix (st, nd, rd, th)
	public static String convertToOrdinal(int number) {
		String words = convert(number);
		if (number % 100 >= 11 && number % 100 <= 13) {
			return words + "TH";
		}
		switch (number % 10) {
		case 1:
			return words + "ST";
		case 2:
			return words + "ND";
		case 3:
			return words + "RD";
		default:
			return words + "TH";
		}
	}
	

	/**
     * Converts a Roman to its words representation.
     */
	public static String convertSessionNumberToWords(String sessionNo) {
    	String sessionInWords=null;
    	
    	if(sessionNo != null)
    	{
        String[] romans = { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX",
                "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX",
                "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXVI", "XXVII", "XXVIII", "XXIX", "XXX" };
        
        String[] words = {
        			    "", // index 0 placeholder
        			    "First", "Second", "Third", "Fourth", "Fifth",
        			    "Sixth", "Seventh", "Eighth", "Ninth", "Tenth",
        			    "Eleventh", "Twelfth", "Thirteenth", "Fourteenth", "Fifteenth",
        			    "Sixteenth", "Seventeenth", "Eighteenth", "Nineteenth", "Twentieth",
        			    "Twenty First", "Twenty Second", "Twenty Third", "Twenty Fourth", "Twenty Fifth",
        			    "Twenty Sixth", "Twenty Seventh", "Twenty Eighth", "Twenty Ninth", "Thirtieth"
        };
        
        for(int i=0;i<romans.length;i++)
        {
        	if(sessionNo.equals(romans[i]))
        		sessionInWords=words[i];
        }
    }
    	return sessionInWords;
    }
    
    /**
     * Converts a lokSabhaNo to its words representation.
     */
    public static String convertLokSabhaNumberToWords(String lokSabhaNo) {
    	String lokSabhaNoInWords=null;
    	
    	if(lokSabhaNo != null)
    	{
        String[] lokSabhaNumbers = { "", "18", "19", "20", "21", "22", "23", "24", "25"};
        
        String[] words = {
        			    "", // index 0 placeholder
        			    "Eighteenth", "Nineteenth", "Twentieth",
        			    "Twenty First", "Twenty Second", "Twenty Third", "Twenty Fourth", "Twenty Fifth"
        };
        
        for(int i=0;i<lokSabhaNumbers.length;i++)
        {
        	if(lokSabhaNo.equals(lokSabhaNumbers[i]))
        		lokSabhaNoInWords=words[i];
        }
    }
    	return lokSabhaNoInWords;
    }
}
