Attribute VB_Name = "modEscapeString"
'---------------------------------------------------------------------------------------
'
'    Copyright 2003 Mike Hillyer (www.vbmysql.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Please forward all questions/suggestions to mike@vbmysql.com
'
'
' Module    : modEscapeString
' DateTime  : 1/2/2004 08:39
' Author    : MIKE HILLYER
' Purpose   : THIS MODULE HOLDS THE CODE NEEDED TO IMPLEMENT MYSQL_ESCAPE_STRING
'             IN A VISUAL BASIC 6 PROJECT. IT SHOULD BE NOTED THAT MYSQL HAS DEPRECIATED
'             MYSQL_ESCAPE_STRING IN FAVOR OF MYSQL_REAL_ESCAPE_STRING BUT SINCE
'             THE LATTER REQUIRES AN ESTABLISHED CONNECTION IT ADDS A LOT OF
'             COMPLEXITY FOR A DEVELOPER WHO IS CURRENTLY USING MYODBC. WE COULD USE
'             MYSQL_REAL_ESCAPE_STRING WHEN USING AN ALL-API APPROACH TO USING MYSQL
'             AND OMITTING MYODBC ENTIRELY.
'                  SEE http://www.mysql.com/doc/en/mysql_escape_string.html
'
'             USER BEWARE: THIS FUNCTION IF FOR USE WITH LATIN(DEFAULT)
'                          CHARACTER SETS ONLY, IT WILL NOT NECESSCARILY
'                          WORK WITH NON-LATIN CHARACTER SETS!
'
'---------------------------------------------------------------------------------------
Option Explicit

'API DECLARATION FOR mysql_escape_string FUNCTION CALL
Public Declare Function api_mysql_escape_string Lib "libmySQL.dll" _
        Alias "mysql_escape_string" _
        (ByVal strTo As String, _
         ByVal strFrom As String, _
         ByVal lngLength As Long _
        ) As Long

Public Function mysql_escape_string(dirtystring As String) As String
Attribute mysql_escape_string.VB_Description = "Calls libmysql.dll mysql_escape_string function to clean a string for insertion into MySQL database. THIS DOES NOT LOOK AT CURRENT DATABASE CHARACTER SET!"
    On Error Resume Next
    Dim strFrom As String           'SOURCE STRING PASSED TO FUNCTION
    Dim lngFromLength As String     'LENGTH OF SOURCE STRING
    Dim strTo As String             'DESTINATION STRING COMING FROM FUNCTION
    Dim lngToLength As Long         'LENGTH OF DESTINATION STRING
    
    strFrom = dirtystring           'STORE FUNCTION INPUT
    lngFromLength = Len(strFrom)    'GET LENGTH OF INPUT
    
    strTo = Space(lngFromLength * 2 + 1) 'ALLOCATE A BUFFER FOR OUTPUT OF FUNCTION
                                         '2 BYTES PER CHARACTER PLUS A BYTE FOR NULL
                                         'TERMINATOR USED BY FUNCTION
    
    lngToLength = api_mysql_escape_string(strTo, strFrom, lngFromLength) 'CALL API
    
    mysql_escape_string = Left(strTo, lngToLength) 'TRIM NULL TERMINATOR
End Function
