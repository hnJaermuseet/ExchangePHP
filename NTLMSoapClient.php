<?php
/**
 * Soap Client using Microsoft's NTLM Authentication
 *
 * Copyright (c) 2010 Jærmuseet http://jaermuseet.no
 * Copyright (c) 2008 Invest-In-France Agency http://www.invest-in-france.org
 *
 * Author : Hallvard Nygård
 * Author : Thomas Rabaix
 *
 * Permission to use, copy, modify, and distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 *
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 *
 * @link http://rabaix.net/en/articles/2008/03/13/using-soap-php-with-ntlm-authentication
 * @author Hallvard Nygård
 * @author Thomas Rabaix
 */

/**
 * Soap Client using Microsoft's NTLM Authentication
 *
 * @link http://rabaix.net/en/articles/2008/03/13/using-soap-php-with-ntlm-authentication
 * @author Hallvard Nygård
 * @author Thomas Rabaix
 */

class NTLMSoapClient extends SoapClient {
	protected $error;
	
	/**
	 * Whether or not to validate ssl certificates
	 *
	 * @var boolean
	 */
	protected $validate = false;
	
	/**
	 * Performs a SOAP request
	 *
	 * @link http://php.net/manual/en/function.soap-soapclient-dorequest.php
	 *
	 * @param string $request the xml soap request
	 * @param string $location the url to request
	 * @param string $action the soap action.
	 * @param integer $version the soap version
	 * @return string the xml soap response.
	 */
	public function __doRequest($request, $location, $action, $version) 
	{
		$headers = array(
			'Method: POST',
			'Connection: Keep-Alive',
			'User-Agent: PHP-SOAP-CURL',
			'Content-Type: text/xml; charset=utf-8',
			'SOAPAction: "'.$action.'"',
		); // end $headers
		$this->__last_request_headers = $headers;
		$ch = curl_init($location);
		
		curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, $this->validate);
		curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, $this->validate);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
		curl_setopt($ch, CURLOPT_POST, true );
		curl_setopt($ch, CURLOPT_POSTFIELDS, $request);
		curl_setopt($ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
		curl_setopt($ch, CURLOPT_USERPWD, $this->_login.':'.$this->_password);
		
		$response = curl_exec($ch);
		if(
			$response === false // Normal curl error
		)
		{
			// Throwing an exception of the curl error
			$txt = 'Curl error: '. curl_error($ch);
			$this->setErrorAndThrowException($txt, $txt, $ch);
		}
		elseif(curl_getinfo($ch, CURLINFO_HTTP_CODE) == '401') // Unauthorized, maybe wrong password
		{
			$this->setErrorAndThrowException('401', '401 Unauthorized', $ch);
		}
		
		curl_close($ch);
		return $response;
	} // end function __doRequest()

	/**
	 * Returns last SOAP request headers
	 *
	 * @link http://php.net/manual/en/function.soap-soapclient-getlastrequestheaders.php
	 *
	 * @return string the last soap request headers
	 */
	public function __getLastRequestHeaders() {
		return implode("n", $this->__last_request_headers)."n";
	} // end function __getLastRequestHeaders()
	
	public function setErrorAndThrowException($error, $exception_txt, $ch = null)
	{
		$this->setError($error);
		if(!is_null($ch))
			curl_close($ch);
		throw new Exception($exception_txt);
	}
	
	public function setError($error)
	{
		$this->error = $error;
	}
	
	public function hasError()
	{
		return isset($this->error);
	}
	
	public function getError()
	{
		if(isset($this->error))
			return $this->error;
		else
			return ''; // No error
	}
	
	/**
	 * Sets whether or not to validate ssl certificates
	 *
	 * @param boolean $validate
	 */
	public function validateCertificate($validate = true) {
		$this->validate = $validate;

		return true;
	} // end public function validateCertificate()
} // end class NTLMSoapClient