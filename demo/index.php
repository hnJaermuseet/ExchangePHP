<?php

/*

DEMO FOR SYNCING BETWEEN PHP&MYSQL AND EXCHANGE 2007
Hallvard Nyg�rd <hn@jaermuseet> for J�rmuseet, 
a Norwegian science center.

J�rmuseet
http://jaermuseet.no/

JM-booking, the project using the ExchangePHP
https://github.com/hnJaermuseet/JM-booking

This demo has
- a set of items with title and text
- For each item a calendar element is created
- Created elements starts now and ends an hour later
- If a calendar element is deleted or changed a new
  one is created
- If an item is changed, the calendar element should
  be deleted and a new element is created

This is mostly based on how JM-booking is going to be syncing
against Exchange.

Thanks to Erik Cederstrand and his article at
http://www.howtoforge.com/talking-soap-with-exchange

He used some code from Thomas Rabaix:
http://rabaix.net/en/articles/2008/03/13/using-soap-php-with-ntlm-authentication

Also, you might want to look at
- http://code.google.com/p/php-ews/

This demo is tested on WAMP with PHP5 and the following
modules turned on:
- php_soap
- php_openssl

In addition to Apache/webserver and PHP5, you must have
- Exchange server (tested against Exchange 2007)
- MySQL-server

Make the database exchangeTest and create these tables:
CREATE TABLE `exchangeTest`.`items` (
	`id` INT NOT NULL AUTO_INCREMENT ,
	`title` VARCHAR( 25 ) NOT NULL ,
	`text` TEXT NOT NULL ,
	PRIMARY KEY ( `id` )
);
CREATE TABLE `exchangeTest`.`sync` (
	`item_id` INT NOT NULL ,
	`e_id` VARCHAR( 255 ) NOT NULL ,
	`e_changekey` VARCHAR( 255 ) NOT NULL ,
	`user` VARCHAR( 255 ) NOT NULL,
	`time` INT NOT NULL
);

Also download Services.wsdl from your EWS and adjust it
like it tells you in
http://www.howtoforge.com/talking-soap-with-exchange
(this also downloads messages.xsd and types.xsd)

*/

require_once dirname(__FILE__).'/../ExchangePHP.php';

function printout ($txt)
{
	global $user_id;
	if(php_sapi_name() == 'cli') // Command line
	{
		echo date('Y-m-d H:i:s').' [user '.$user_id.'] '.$txt."\r\n";
	}
	else
	{
		echo str_replace(' ', '&nbsp;', 
			date('Y-m-d H:i:s').' [user '.$user_id.'] '.$txt).'<br />'.chr(10);
	}
}

// Exchange login
require dirname(__FILE__).'/password.php';
/* Syntax:
$login = array(
		'username' => '',
		'password' => '',
	);
*/

// MySQL
mysql_connect('localhost', 'exchangetest', '');
mysql_select_db('exchangeTest');

if(isset($_GET['erase']))
{
	mysql_query("TRUNCATE `items`;");
	echo mysql_error();
	mysql_query("TRUNCATE `sync`;");
	echo mysql_error();
	mysql_query("INSERT INTO `exchangeTest`.`items` (
		`id` ,
		`title` ,
		`text`
	)
	VALUES (
		NULL , 'A', 'lalala'
	), (
		NULL , 'B', 'hello'
	), (
		NULL , 'C', 'yeah'
	);");
	echo mysql_error();
	
	printout('DB reset');
	exit;
}

if(php_sapi_name() != 'cli')
{
	echo '<a href="">Run</a> -:- ';
	echo '<a href="?erase=1">Reset database</a>';
	echo '<br /><br />'.chr(10).chr(10);
}

// Getting items
$items = array();
$Q = mysql_query("select * from `items`");
echo mysql_error();
while($R_item = mysql_fetch_assoc($Q))
{
	$items[$R_item['id']] = $R_item;
}

// Getting sync-data
$sync = array();
$Q = mysql_query("select * from `sync` where `user` = '".$login['username']."'");
echo mysql_error();
while($R_sync = mysql_fetch_assoc($Q))
{
	$sync[$R_sync['item_id']] = $R_sync;
}

$wsdl = dirname(__FILE__).'/Services.wsdl';
$client = new NTLMSoapClient($wsdl, array(
		'login'       => $login['username'], 
		'password'    => $login['password'],
		'trace'       => true,
		'exceptions'  => true,
	)); 

$cal = new ExchangePHP($client);

try
{
	$calendaritems = $cal->getCalendarItems(
		date('Y-m-d').'T00:00:00', // Today
		date('Y-m-d',time()+61171200).'T00:00:00Z' // Approx 2 years, seems to be a limit
	);
}
catch (Exception $e)
{
	printout('Exception - getCalendarItems: '.$e->getMessage());
	
	if($cal->client->getError() == '401')
	{
		// Unauthorized
		printout('Wrong username and password.');
	}
	exit;
}


$cal_ids = array(); // Id => ChangeKey
if(is_null($calendaritems))
{
	printout('getCalendarItems failed: '.$cal->getError());
}
else
{
	// Going through existing elements
	foreach($calendaritems as $item) {
		if(!isset($item->Subject))
			$item->Subject = '';
		$cal_ids[$item->ItemId->Id] = $item->ItemId->ChangeKey;
		printout('Existing: '.$item->Start.'   '.$item->End.'   '.$item->Subject);
	}
}

// Analysing which to create
$items_new     = array();
$items_delete  = array();
foreach($items as $item) // Running through items in database
{
	// Checking for previous sync
	$delete = false;
	if(isset($sync[$item['id']]))
	{
		// Check if it is deleted in Exchange
		$this_sync = $sync[$item['id']];
		if(!isset($cal_ids[$this_sync['e_id']]))
		{
			printout('<span color="red">Err! Calendar element is deleted in Exchange!</span>');
			$create_new = true;
		}
		else
		{
			// Check if it is changed in Exchange
			if($cal_ids[$this_sync['e_id']] != $this_sync['e_changekey'])
			{
				printout('<span color="red">Err! Calendar element is changed in Exchange!</span>');
				$create_new = true;
			}
			else
			{
				// No changes in Exchange
				$create_new = false;
				
				// Is there any new revisions of the item-data?
				// => trigger delete and create a new
				if($this_sync['time'] < time()+60) // Using a minute as change
				{
					$create_new  = true;
					$delete      = true;
					$delete_id   = $cal_ids[$this_sync['e_id']];
				}
			}
		}
	}
	else
	{
		// Never synced before, create a new one
		$create_new = true;
	}
	
	if($create_new)
	{
		$i = $cal->createCalendarItems_addItem($item['title'], $item['text'], date('c'), date('c', time()+60*60),
				array(
					'ReminderIsSet' => false,
					'Location' => 'Interwebs',
				)
			);
		$items_new[$i] = $item['id'];
	}
	if($delete)
	{
		$items_delete[$item['id']] = $item;
	}
}

/* ## CREATE ITEMS ## */
try
{
	$created_items = $cal->createCalendarItems();
}
catch (Exception $e)
{
	printout('Exception - createCalendarItems: '.$e->getMessage().'<br />');
	$created_items = array();
}

foreach($created_items as $i => $ids)
{
	if(!is_null($ids['Id'])) // Null = unsuccessful
	{
		printout($items_new[$i].' created.');
		// Deleting from sync
		mysql_query("DELETE FROM `sync`
			WHERE
				`item_id` = '".$items_new[$i]."'
			");
		echo mysql_error();
		// Inserting in sync
		mysql_query("INSERT INTO `sync` (
			`item_id` ,
			`e_id` ,
			`e_changekey`,
			`user`,
			`time`
		)
		VALUES (
			'".$items_new[$i]."' , 
			'".$ids['Id']."', 
			'".$ids['ChangeKey']."',
			'".$login['username']."',
			'".time()."'
		);");
		echo mysql_error();
	}
	else
	{
		printout($items_new[$user_id][$i] .' not created: '.print_r($ids['ResponseMessage'], true));
	}
}

/* ## DELETE ITEMS ## */
$deleted_items = array();
foreach($items_delete as $item_id => $item)
{
	try
	{
		$deleted_item = $cal->deleteItem($sync[$item['id']]['e_id']);
		$deleted_items[$item['id']] = $deleted_item;
		printout($item['id'].' deleted');
	}
	catch (Exception $e)
	{
		printout('Exception - deleteItem - '.$item['id'].': '.$e->getMessage());
	}
}
/*
foreach($deleted_items as $i => $ids)
{
	if(!is_null($ids)) // Null = unsuccessful
	{
		printout($items_new[$i].' synced (created).');
		mysql_query("INSERT INTO `sync` (
			`item_id` ,
			`e_id` ,
			`e_changekey`,
			`user`,
			`time`
		)
		VALUES (
			'".$items_new[$i]."' , 
			'".$ids['Id']."', 
			'".$ids['ChangeKey']."',
			'".$login['username']."',
			'".time()."'
		);");
		echo mysql_error();
	}
}*/