<?php

/*

DEMO FOR SYNCING BETWEEN PHP&MYSQL AND EXCHANGE 2007
Hallvard Nygård <hn@jaermuseet> for Jærmuseet, 
a Norwegian science center.

Jærmuseet
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

require_once '../ExchangePHP.php';

// Exchange login
$login = array(
		'username' => 'hn',
		'password' => 'Nsar01',
	);

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
	
	echo 'DB reset';
	exit;
}
echo '<a href="">Run</a> -:- ';
echo '<a href="?erase=1">Reset database</a>';
echo '<br /><br />';

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

stream_wrapper_unregister('https');
stream_wrapper_register('https', 'NTLMStream') or die("Failed to register protocol");

$wsdl = "Services.wsdl";
$client = new NTLMSoapClient($wsdl, array(
		'login'     => $login['username'], 
		'password'  => $login['password'],
		'trace'     => 1,
	)); 

/* Do something with the web service connection */
stream_wrapper_restore('https');


$cal = new ExchangePHP($client);
try
{
	$calendaritems = $cal->getCalendarItems("2010-10-01T00:00:00Z","2010-12-31T00:00:00Z");
}
catch (Exception $e)
{
	echo 'Exception: '.$e->getMessage().'<br />';
}

// Going through existing elements
$cal_ids = array(); // Id => ChangeKey
foreach($calendaritems as $item) {
	if(!isset($item->Subject))
		$item->Subject = '';
	$cal_ids[$item->ItemId->Id] = $item->ItemId->ChangeKey;
	echo 'Existing: '.$item->Start.'&nbsp;&nbsp;&nbsp;'.$item->End.'&nbsp;&nbsp;&nbsp;'.$item->Subject."<br>";
} 

// Analysing which to create
$items_new     = array();
$items_delete  = array();
foreach($items as $item)
{
	// Checking for previous sync
	$delete = false;
	if(isset($sync[$item['id']]))
	{
		// Check if it is deleted in Exchange
		$this_sync = $sync[$item['id']];
		if(!isset($cal_ids[$this_sync['e_id']]))
		{
			echo '<span color="red">Err! Calendar element is deleted in Exchange!</span><br />';
			$create_new = true;
		}
		else
		{
			// Check if it is changed in Exchange
			if($cal_ids[$this_sync['e_id']] != $this_sync['e_changekey'])
			{
				echo '<span color="red">Err! Calendar element is changed in Exchange!</span><br />';
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
		$i = $cal->createItems_addItem($item['title'], $item['text'], date('c'), date('c', time()+60*60),
				array(
					'ReminderIsSet' => false,
					'Location' => 'Interwebs',
				)
			);
		$items_new[$i] = $item['id'];
	}
	if($delete)
	{
		//$i = $cal->deleteItem_addItem($delete_id['Id'], $delete_id['changeKey']);
		//$items_delete[$i] = $item['id'];
	}
}

try
{
	$created_items = $cal->createItems();
}
catch (Exception $e)
{
	echo 'Exception: '.$e->getMessage().'<br />';
	$created_items = array();
}

foreach($created_items as $i => $ids)
{
	if(!is_null($ids)) // Null = unsuccessful
	{
		echo $items_new[$i].' synced (created).<br />';
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
}