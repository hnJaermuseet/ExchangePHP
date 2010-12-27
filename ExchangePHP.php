<?php

require_once dirname(__FILE__).'/NTLMSoapClient.php';
require_once dirname(__FILE__).'/NTLMStream.php';

/**
 * EXCHANGEPHP
 * PHP library for talking to Exchange 2007 with SOAP
 * Made by Hallvard Nygård <hn@jaermuseet> for use in
 * JM-booking. JM-booking is a booking system used by
 * Jærmuseet, a Norwegian science center
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
	public $have_started_deleteItem = false;
	
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
	 * @return  Array    Items or null
	 */
	public function getCalendarItems($from, $to)
	{
		$FindItem->Traversal = "Shallow";
		$FindItem->ItemShape->BaseShape = "AllProperties";
		$FindItem->ParentFolderIds->DistinguishedFolderId->Id = "calendar";
		$FindItem->CalendarView->StartDate = $from;
		$FindItem->CalendarView->EndDate = $to;
		$result = $this->client->FindItem($FindItem);
		if($result->ResponseMessages->FindItemResponseMessage->ResponseClass != 'Success')
		{
			$this->setError($result->ResponseMessages->FindItemResponseMessage->MessageText);
			return null;
		}
		else
			return $result->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem;
	}
	
	/**
	 * Makes the CreateItem object with some default variables
	 * Used by addCalendarItem
	 */
	protected function createCalendarItems_startup ()
	{
		$this->CreateItem->SavedItemFolderId->DistinguishedFolderId->Id = 'calendar';
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
	 * @return  int     Internal number identifying the item
	 */
	public function createCalendarItems_addItem($title, $text, $start, $end, $options)
	{
		if(!$this->have_started_createItem)
			$this->createCalendarItems_startup();
		if(!is_array($options))
			throw new Exception('Options must be array');
		
		$i = count($this->CreateItem->Items->CalendarItem);
		
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
			throw new Exception ('No items added.');
		if(count($this->CreateItem->Items->CalendarItem) == 0)
			throw new Exception ('No items added.');
		
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
					$ids[$i] = null;
				}
			}
			return $ids;
		}
		
		if ( $result->ResponseMessages->CreateItemResponseMessage->ResponseClass == 'Success' )
		{
			// Returning the id/changekey in an array with only one item
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
			return array(0 => null);
		}
	}
	
	/**
	 * Mkes the DeleteItem object with some default variables
	 * Used by addCalendarItem
	 */
	protected function deleteCalendarItems_startup ()
	{
		//$this->DeleteItem->DeleteType = 'HardDelete';
		$this->DeleteItem->DeleteType = 'SoftDelete';
		$this->DeleteItem->ItemIds = array();
		
		$this->have_started_deleteItem = true;
	}
	
	/**
	 * Adds an items to the list of items to be deleted
	 *
	 * @param   string  Exchange ID
	 * @return  int     Internal number identifying the item
	 */
	public function deleteItems_addItem($id)
	{
		if(!$this->have_started_deleteItem)
			$this->deleteCalendarItems_startup();
		
		$i = count($this->DeleteItem->ItemIds);
		$this->DeleteItem->ItemIds[$i] = array('ItemId' => $id);
		
		return $i;
	}
	
	/**
	 * Deletes the items added by deleteItems_addItem
	 *
	 * @return  array  Response message(s)
	 */
	public function deleteItem($id)//, $changekey)
	{
		//if(!$this->have_started_deleteItem)
		//	throw new Exception ('No items added.');
		//if(count($this->DeleteItem->ItemIds) == 0)
		//	throw new Exception ('No items added.');
		
		$this->DeleteItem->DeleteType = 'SoftDelete';
		$this->DeleteItem->ItemIds->ItemId->Id = $id; /* = array(
				'Id' => $id,
				//'ChangeKey' => $changekey,
			);/**/
		$this->DeleteItem->SendMeetingCancellations = 'SendToNone';
		
		$result = $this->client->DeleteItem($this->DeleteItem); // < $this->client holds SOAP-client object
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
}