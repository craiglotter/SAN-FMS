<?php

function sort_ldap_entries($e, $fld, $order)
{
   for ($i = 0; $i < $e['count']; $i++) {
       for ($j = $i; $j < $e['count']; $j++) {
           $d = strcasecmp($e[$i][$fld][0], $e[$j][$fld][0]);
           switch ($order) {
           case 'A':
               if ($d > 0)
                   swap($e, $i, $j);
               break;
           case 'D':
               if ($d < 0)
                   swap($e, $i, $j);
               break;
           }
       }
   }
   return ($e);
}

function swap(&$ary, $i, $j)
{
   $temp = $ary[$i];
   $ary[$i] = $ary[$j];
   $ary[$j] = $temp;
}


$ldap_server = "ldap://127.0.0.1:389";
$connect=@ldap_connect($ldap_server);
ldap_set_option($connect, LDAP_OPT_PROTOCOL_VERSION, 3);
$filter="(&(cn=" . $_GET["group"] . ")(objectclass=groupOfNames))"; 
$justthese = array( "cn","member");
$sr=ldap_search($connect, "ou=com,ou=main,o=uct", $filter ,$justthese); 



           ldap_sort($connect, $sr, "cn");

		
				$info = ldap_get_entries($connect, $sr);

   				$entry = ldap_first_entry($connect, $sr);
   			
 
	for ($i=0; $i<$info["count"]; $i++) 
   				{
					
						if (isset($info[$i]["member"]))
						{
						sort($info[$i]["member"]);
						}
						for ($j=0; $j < (count($info[$i]["member"])-1); $j++) 
						{
						echo "" . $info[$i]["member"][$j] . "\n";
						}
   				$entry = ldap_next_entry($connect, $entry);					
				}
@ldap_close($connect);
?>
