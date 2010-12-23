<?php

require_once 'NTLMSoapClient.php';
require_once 'NTLMStream.php';

/**
 * EXCHANGEPHP
 * PHP library for talking to Exchange 2007 with SOAP
 * Made by Hallvard Nyg�rd <hn@jaermuseet> for use in
 * JM-booking. JM-booking is a booking system used by
 * J�rmuseet, a Norwegian science center
 *
 * CC-BY 2.0
 *
 * See demo for how to set this thing up and how to use
 * it.
 * 
 * J�rmuseet
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
	protected $client;
	
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
			return null;
		else
			return $result->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem;
	}
	
	/**
	 * Makes the CreateItem object with some default variables
	 * Used by addCalendarItem
	 */
	protected function createItems_startup ()
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
	public function createItems_addItem($title, $text, $start, $end, $options)
	{
		if(!$this->have_started_createItem)
			$this->createItems_startup();
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
	 * Creates the items added by addCalendarItem
	 *
	 * @return  array  Response message(s)
	 */
	public function createItems()
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
}