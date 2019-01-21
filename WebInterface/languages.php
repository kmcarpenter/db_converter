<?php
	// ***************************************************************************
	// *  Copyright 2002-2005 Michael Carpenter and Zenwerx Custom Programming   *
	// ***************************************************************************
	// *                                                                         *
	// *  Mailing Address:                                                       *
	// *                                                                         *
	// *  Zenwerx Custom Programming                                             *
	// *  c/o Michael Carpenter                                                  *
	// *  10 Madison Ave                                                         *
	// *  Brantford , Ontario, Canada                                            *
	// *  N3T 5X3                                                                *
	// *                                                                         *
	// ***************************************************************************
	// *                                                                         *
	// *  Email Address:                                                         *
	// *                                                                         *
	// *  zenwerx@zenwerx.com                                                    *
	// *                                                                         *
	// ***************************************************************************
	// *                                                                         *
	// *  Web Address:                                                           *
	// *                                                                         *
	// *  http://www.zenwerx.com                                                 *
	// *                                                                         *
	// ***************************************************************************
	//
	//    This file is part of DB Converter 1.6.0.0
	//
	//    DB Converter 1.6.0.0 is free software; you can redistribute it and/or
	//    modify it under the terms of the GNU General Public License as published by
	//    the Free Software Foundation; either version 2 of the License, or
	//    (at your option) any later version.
	//
	//    DB Converter 1.6.0.0 is distributed in the hope that it will be useful,
	//    but WITHOUT ANY WARRANTY; without even the implied warranty of
	//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	//    GNU General Public License for more details.'
	//
	//   You should have received a copy of the GNU General Public License
	//    along with DB Converter 1.6.0.0; if not, write to the Free Software
	//    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

    require( "defines.inc.php"      );
    require( "functions.inc.php"    );

    define( "ERR_STRING", "0"     );
    if (!($link=connectDB()))
            echo ERR_STRING; // default to correct version

    if ($_GET['action']=='0')
    {
            $query = "SELECT lang_id, lang_name FROM dbc_languages WHERE version=\"".$_GET['version']."\"";

            if (!($result = mysql_query($query)))
            {
                    echo ERR_STRING;
            } else
            {
                    $c = mysql_num_rows($result);
                    for ($i=0;$i<$c;$i++)
                    {
                            $row = mysql_fetch_row($result);
                            echo $row[0] . "," . $row[1] . "|";
                    }
            }

    } else if ($_GET['action']=='1')
    {
        $query = "SELECT lang_file, lang_size, version FROM dbc_languages WHERE lang_id=" . $_GET['lang'];

		if (!($result = mysql_query($query)))
		{
			echo ERR_STRING;
		} else
		{
			$row = mysql_fetch_row($result);
			echo "http://". $_SERVER['HTTP_HOST'] . "/languages/" . $row[2] . "/" . $row[0] ."," . $row[1]; 
		}
    }
?>
