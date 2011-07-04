<?php

require_once dirname(__FILE__).'/NTLMSoapClient.php';

/**
 * EXCHANGEPHP
 * PHP library for talking to Exchange 2007 with SOAP
 * Made by Hallvard Nygård <hn@jaermuseet> for use in
 * JM-booking. JM-booking is a booking system used by
 * Jærmuseet, a Norwegian museum and science center
 *
 * CC-BY 2.0
 *
 * See demo for how to set this thing up and how to use
 * it.
 * 
 * Jærmuseet
 * http://jaermuseet.no/
 * 
 * JM-booking, the project using the ExchangePHP
 * https://github.com/hnJaermuseet/JM-booking
 *
 * Use of FindItem is based on Erik Cederstrands article:
 * http://www.howtoforge.com/talking-soap-with-exchange
 * 
 * The creation of the "CreateItem" is based on:
 * http://www.howtoforge.com/forums/showpost.php?p=226530&postcount=28
 * 
 * He used some code from Thomas Rabaix:
 * http://rabaix.net/en/articles/2008/03/13/using-soap-php-with-ntlm-authentication
 * 
 * Also, you might want to look at
 * http://code.google.com/p/php-ews/
 */

class ExchangePHP
{
	public $client;
	protected $error;
	protected $responseClass = '';
	protected $responseCode = '';
	
	protected $calendaritem_valid_options =
		array(
			'ReminderIsSet',
			'ReminderMinutesBeforeStart',
			'IsAllDayEvent',
			'LegacyFreeBusyStatus',
			'Location',
			// TODO: Add more
		);
	
	public $have_started_createItem = false;
	
	public function __construct ($client)
	{
		$this->client = $client;
	}
	
	/**
	 * Gets all calendar items in a time period and returns them
	 * 
	 * Based on Erik Cederstrands article:
	 * http://www.howtoforge.com/talking-soap-with-exchange
	 *
	 * @param   String   From, ISO date
	 * @param   String   To, ISO date
	 * @param   String   Optional, what users calendar to get from. Must be primary STMP email address
	 * @return  Array    Items or null
	 */
	public function getCalendarItems($from, $to, $primary_emailaddress = null)
	{
		// Init
		$FindItem = new stdClass();
		$FindItem->ItemShape = new stdClass();
		$FindItem->ParentFolderIds = new stdClass();
		$FindItem->ParentFolderIds->DistinguishedFolderId = new stdClass();
		$FindItem->CalendarView = new stdClass();
		$FindItem->ParentFolderIds->DistinguishedFolderId->Mailbox = new stdClass();
		
		
		$FindItem->Traversal = "Shallow";
		$FindItem->ItemShape->BaseShape = "AllProperties";
		$FindItem->ParentFolderIds->DistinguishedFolderId->Id = "calendar";
		if(!is_null($primary_emailaddress))
			$FindItem->ParentFolderIds->DistinguishedFolderId->Mailbox->EmailAddress = $primary_emailaddress;
		$FindItem->CalendarView->StartDate = $from;
		$FindItem->CalendarView->EndDate = $to;
		$result = $this->client->FindItem($FindItem);
		
		$this->setResponseCode($result->ResponseMessages->FindItemResponseMessage->ResponseCode);
		$this->setResponseClass($result->ResponseMessages->FindItemResponseMessage->ResponseClass);
		if($result->ResponseMessages->FindItemResponseMessage->ResponseClass != 'Success')
		{
			$this->setError($result->ResponseMessages->FindItemResponseMessage->MessageText);
			return null;
		}
		else
		{
			if($result->ResponseMessages->FindItemResponseMessage->RootFolder->TotalItemsInView > 1)
				return $result->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem;
			elseif($result->ResponseMessages->FindItemResponseMessage->RootFolder->TotalItemsInView == 1)
				return array($result->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem);
			else
				return array();
		}
	}
	
	/**
	 * Makes the CreateItem object with some default variables
	 * Used by addCalendarItem
	 *
	 * @param   String  Optional, what users calendar to get from. Must be primary STMP email address
	 */
	protected function createCalendarItems_startup ($primary_emailaddress = null)
	{
		$this->CreateItem = new stdClass();
		$this->CreateItem->SavedItemFolderId = new stdClass();
		$this->CreateItem->SavedItemFolderId->DistinguishedFolderId = new stdClass();
		$this->CreateItem->SavedItemFolderId->DistinguishedFolderId->Mailbox = new stdClass();
		$this->CreateItem->Items = new stdClass();
		
		$this->CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = 'calendar';
		if(!is_null($primary_emailaddress))
		{
			$this->CreateItem->SavedItemFolderId->DistinguishedFolderId->Mailbox->EmailAddress = $primary_emailaddress;
		}
		$this->CreateItem->Items->CalendarItem = array();
		$this->CreateItem->SendMeetingInvitations = "SendToNone";
		
		$this->have_started_createItem = true;
	}
	
	/**
	 * Adds an items to the list of items to be created
	 *
	 * @param   string  Title of the calendar item
	 * @param   string  Text body of the calendar item
	 * @param   string  Start time, ISO format
	 * @param   string  End time, ISO format
	 * @param   array   Options that are allowed to the to Exchange
	 * @param   String  Optional, what users calendar to get from. Must be primary STMP email address
	 * @return  int     Internal number identifying the item
	 */
	public function createCalendarItems_addItem($title, $text, $start, $end, $options, $primary_emailaddress = null)
	{
		if(!$this->have_started_createItem)
			$this->createCalendarItems_startup($primary_emailaddress);
		if(!is_array($options))
			throw new Exception('Options must be array');
		
		$i = count($this->CreateItem->Items->CalendarItem);
		
		$this->CreateItem->Items->CalendarItem[$i] = new stdClass();
		
		$this->CreateItem->Items->CalendarItem[$i]->Subject = $title;
		
		$this->CreateItem->Items->CalendarItem[$i]->Start = $start; # ISO date format. Z denotes UTC time
		$this->CreateItem->Items->CalendarItem[$i]->End = $end;
		
		$this->CreateItem->Items->CalendarItem[$i]->Body['_'] = $text; 
		$this->CreateItem->Items->CalendarItem[$i]->Body['BodyType']="Text";
		
		foreach($options as $option => $value)
		{
			if(in_array($option,$this->calendaritem_valid_options))
			{
				$this->CreateItem->Items->CalendarItem[$i]->{$option} = $value;
			}
		}
		
		//$this->CreateItem->Items->CalendarItem[$i]->RequiredAttendees->Attendee[]->Mailbox->EmailAddress = 'hn@jaermuseet.no';
		//$this->CreateItem->Items->CalendarItem[$i]->RequiredAttendees->Attendee[]->Mailbox->EmailAddress = 'runar.sandsmark@jaermuseet.no';
		
		return $i;
	}
	
	/**
	 * Creates the items added by createCalendarItems_addItem
	 *
	 * @return  array  Response message(s)
	 */
	public function createCalendarItems()
	{
		if(!$this->have_started_createItem)
		{
			$this->createCalenderItems_shutdown();
			throw new Exception ('No items added.');
		}
		if(count($this->CreateItem->Items->CalendarItem) == 0)
		{
			$this->createCalenderItems_shutdown();
			throw new Exception ('No items added.');
		}
		
		$result = $this->client->CreateItem($this->CreateItem); // < $this->client holds SOAP-client object
		if(count($this->CreateItem->Items->CalendarItem) > 1)
		{
			$ids = array();
			foreach($result->ResponseMessages->CreateItemResponseMessage as $i => $response)
			{
				if($response->ResponseClass == 'Success')
				{
					// Returning the id/changekey for each item
					$ids[$i] = array(
						'Id' =>  $response->Items->CalendarItem->ItemId->Id,
						'ChangeKey' => $response->Items->CalendarItem->ItemId->ChangeKey,
					);
				}
				else
				{
					// If there is no success, return null
					printout('Item '.$i.' failed creation: '.print_r($response, true));
					$ids[$i] = array(
							'Id' => null,
							'ResponseMessage' => $response
						);
				}
			}
			
			$this->createCalenderItems_shutdown();
			return $ids;
		}
		
		if ( $result->ResponseMessages->CreateItemResponseMessage->ResponseClass == 'Success' )
		{
			// Returning the id/changekey in an array with only one item
			$this->createCalenderItems_shutdown();
			return array(0 => 
				array(
					'Id' =>  $result->ResponseMessages->CreateItemResponseMessage->Items->CalendarItem->ItemId->Id,
					'ChangeKey' => $result->ResponseMessages->CreateItemResponseMessage->Items->CalendarItem->ItemId->ChangeKey,
				)
			);
		}
		else
		{
			// Return null
			$this->createCalenderItems_shutdown();
			return array(0 => array(
					'Id' => null,
					'ResponseMessage' => $result->ResponseMessages->CreateItemResponseMessage
				));
		}
	}
	
	/**
	 * After creation of calendar items (or it has failed)
	 * we'll do some clean up
	 */
	protected function createCalenderItems_shutdown()
	{
		unset($this->CreateItem);
		$this->have_started_createItem = false;
	}
	
	/**
	 * Delete the item
	 *
	 * @param   String  id of the item
	 * @return  array   Response message(s)
	 */
	public function deleteItem($id)//, $changekey)
	{
		$this->DeleteItem = new stdClass();
		$this->DeleteItem->ItemIds = new stdClass();
		$this->DeleteItem->ItemIds->ItemId = new stdClass();
		
		$this->DeleteItem->DeleteType = 'SoftDelete';
		$this->DeleteItem->ItemIds->ItemId->Id = $id; /* = array(
				'Id' => $id,
				//'ChangeKey' => $changekey,
			);/**/
		$this->DeleteItem->SendMeetingCancellations = 'SendToNone';
		
		$result = $this->client->DeleteItem($this->DeleteItem); // < $this->client holds SOAP-client object
		
		unset($this->DeleteItem); // Clean up
		return ($result->ResponseMessages->DeleteItemResponseMessage->ResponseClass == 'Success');
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
	
	public function setResponseClass($a) {
		$this->responseClass = $a;
	}
	public function getResponseClass() {
		return $this->responseClass;
	}
	public function setResponseCode($a) {
		$this->responseCode = $a;
	}
	public function getResponseCode() {
		return $this->responseCode;
	}
}