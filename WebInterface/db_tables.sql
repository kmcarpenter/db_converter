-- ***************************************************************************
-- *  Copyright 2002-2005 Michael Carpenter and Zenwerx Custom Programming   *
-- ***************************************************************************
-- *                                                                         *
-- *  Mailing Address:                                                       *
-- *                                                                         *
-- *  Zenwerx Custom Programming                                             *
-- *  c/o Michael Carpenter                                                  *
-- *  10 Madison Ave                                                         *
-- *  Brantford , Ontario, Canada                                            *
-- *  N3T 5X3                                                                *
-- *                                                                         *
-- ***************************************************************************
-- *                                                                         *
-- *  Email Address:                                                         *
-- *                                                                         *
-- *  zenwerx@zenwerx.com                                                    *
-- *                                                                         *
-- ***************************************************************************
-- *                                                                         *
-- *  Web Address:                                                           *
-- *                                                                         *
-- *  http://www.zenwerx.com                                                 *
-- *                                                                         *
-- ***************************************************************************
--
--    This file is part of DB Converter 1.6.0.0
--
--    DB Converter 1.6.0.0 is free software; you can redistribute it and/or
--    modify it under the terms of the GNU General Public License as published by
--    the Free Software Foundation; either version 2 of the License, or
--    (at your option) any later version.
--
--    DB Converter 1.6.0.0 is distributed in the hope that it will be useful,
--    but WITHOUT ANY WARRANTY; without even the implied warranty of
--    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
--    GNU General Public License for more details.'
--
--   You should have received a copy of the GNU General Public License
--    along with DB Converter 1.6.0.0; if not, write to the Free Software
--    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


-- 
-- Table structure for table `dbc_languages`
-- 

CREATE TABLE `dbc_languages` (
  `lang_id` bigint(20) unsigned NOT NULL auto_increment,
  `lang_name` varchar(50) NOT NULL default '',
  `lang_file` varchar(50) NOT NULL default '',
  `lang_size` bigint(20) NOT NULL default '0',
  `version` varchar(10) NOT NULL default '',
  PRIMARY KEY  (`lang_id`)
) TYPE=MyISAM;

-- 
-- Dumping data for table `dbc_languages`
-- 

INSERT INTO `dbc_languages` (`lang_id`, `lang_name`, `lang_file`, `lang_size`, `version`) VALUES (1, 'English', 'EnglishPack.zip', 664517, '1.6.1');
INSERT INTO `dbc_languages` (`lang_id`, `lang_name`, `lang_file`, `lang_size`, `version`) VALUES (2, 'Español', 'SpanishPack.zip', 5372, '1.6.1');

-- --------------------------------------------------------

-- 
-- Table structure for table `dbc_version`
-- 

CREATE TABLE `dbc_version` (
  `CurrentVersion` char(7) default NULL,
  `fSize` bigint(20) default NULL
) TYPE=MyISAM;

-- 
-- Dumping data for table `dbc_version`
-- 

INSERT INTO `dbc_version` (`CurrentVersion`, `fSize`) VALUES ('1.6.1.0', 4424704);