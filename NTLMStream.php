<?php

class NTLMStream
{
	private $path;
	private $mode;
	private $options;
	private $opened_path;
	private $buffer;
	private $pos;
	
	public function stream_open($path, $mode, $options, $opened_path)
	{
		//echo "[NTLMStream::stream_open] $path , mode=$mode n<br />";
		$this->path = $path;
		$this->mode = $mode;
		$this->options = $options;
		$this->opened_path = $opened_path;
		$this->createBuffer($path);
		return true;
	}
	public function stream_close()
	{
		//echo "[NTLMStream::stream_close] n<br />";
		curl_close($this->ch);
	}
	public function stream_read($count)
	{
		//echo "[NTLMStream::stream_read] $count n<br />";
		if(strlen($this->buffer) == 0)
		{
			return false;
		}
		$read = substr($this->buffer,$this->pos, $count);
		$this->pos += $count;
		return $read;
	}

	public function stream_write($data)
	{
		//echo "[NTLMStream::stream_write] n<br />";
		if(strlen($this->buffer) == 0)
		{
			return false;
		}
		return true;
	}
	public function stream_eof()
	{
		//echo "[NTLMStream::stream_eof] ";
		if($this->pos > strlen($this->buffer)) {
			//echo "true n<br />";
			return true;
		}
		//echo "false n<br />";
		return false;
	}
	/* return the position of the current read pointer */
	public function stream_tell()
	{
		//echo "[NTLMStream::stream_tell] n<br />";
		return $this->pos;
	}
	public function stream_flush()
	{
		//echo "[NTLMStream::stream_flush] n<br />";
		$this->buffer = null;
		$this->pos = null;
	}
	public function stream_stat() 
	{
		//echo "[NTLMStream::stream_stat] n<br />";
		$this->createBuffer($this->path);
		$stat = array( 'size' => strlen($this->buffer), );
		return $stat;
	}
	public function url_stat($path, $flags)
	{
		//echo "[NTLMStream::url_stat] n<br />";
		$this->createBuffer($path);
		$stat = array( 'size' => strlen($this->buffer), );
		return $stat;
	}

	/* Create the buffer by requesting the url through cURL */
	private function createBuffer($path)
	{
		if($this->buffer)
		{
			return;
		}
		//echo "[NTLMStream::createBuffer] create buffer from : $path<br />";
		$this->ch = curl_init($path);
		curl_setopt($this->ch, CURLOPT_RETURNTRANSFER, true);
		curl_setopt($this->ch, CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_1_1);
		curl_setopt($this->ch, CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
		curl_setopt($this->ch, CURLOPT_USERPWD, $this->user.':'.$this->password);
		
		curl_setopt($this->ch, CURLOPT_SSL_VERIFYPEER, false);
		curl_setopt($this->ch, CURLOPT_SSL_VERIFYHOST, false); 
		
		$action ='';
		$headers = array( 
			'Method: POST', 
			'Connection: Keep-Alive', 
			'User-Agent: PHP-SOAP-CURL', 
			'Content-Type: text/xml; charset=utf-8', 
			'SOAPAction: "'.$action.'"', ); 
		curl_setopt($this->ch, CURLOPT_HTTPHEADER, $headers); 
		curl_setopt($this->ch, CURLOPT_POST, true ); 
		//curl_setopt($this->ch, CURLOPT_POSTFIELDS, $request); 
		
		
		
		//echo $this->buffer = curl_exec($this->ch);
		//var_dump($this->buffer);
		//echo "[NTLMStream::createBuffer] buffer size : ".strlen($this->buffer)."bytesn<br />";
		$this->pos = 0;
	}
}