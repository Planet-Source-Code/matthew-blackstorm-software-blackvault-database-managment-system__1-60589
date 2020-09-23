Attribute VB_Name = "mdlError"
Option Explicit

Public Function DisplayError(ErrorNum As Integer) As String
Select Case ErrorNum
    Case 1001
        
    Case 2001
        DisplayError = "SQL statment invalid, is not a valid querry type ie. SELECT, DELETE, UPDATE, APPEND"
    Case 2002
        DisplayError = "SQL statment invalid, no Table to select from"
    Case 2003
        DisplayError = "SQL statment invalid, querry over too many Tables, two related Tables only"
    Case 2004
        DisplayError = "SQL statment invalid, the Table does not exist"
    Case 2005
        DisplayError = "SQL statment invalid, no realtionship exists between Tables"
    Case 2006
        DisplayError = "SQL statment invalid, one or more Tables do not exist from statment"
    Case 2007
        DisplayError = "SQL statment invalid, too many opening brackets in Select statment for a Table or Feild ie. [[ ]"
    Case 2008
        DisplayError = "SQL statment invalid, too many opening brackets in Select statment for a Table or Feild ie. [ ]]"
    Case 2009
        DisplayError = "SQL Statment invalid, Feild from Table in Select statment does not exist"
    Case 2010
        DisplayError = "SQL statment invalid, Table in Select statment does not exist"
    Case 2011
        DisplayError = "SQL statment invalid, too many closing brackets in Select statment ie. ( ))"
    Case 2012
        DisplayError = "SQL statment invalid, too many opening brackets in Select statment ie. (( )"
    Case 2013
        DisplayError = "SQL statment invalid, Table or Feild name not completed in Select statment"
    Case 2014
        DisplayError = "SQL statment invalid, all quotes are not finished in Where Statment"
    Case 2015
        DisplayError = "SQL statment invalid, too many opening brackets in Where statment for a Table or Feild ie. [[ ]"
    Case 2016
        DisplayError = "SQL statment invalid, too many closing brackets in Where statment for a Table or Feild ie. [ ]]"
    Case 2017
        DisplayError = "SQL statment invalid, must use a '!' to seperate a Table and Feild in Select statment ie. [Temp]![A]"
    Case 2018
        DisplayError = "SQL statment invalid, must use a '!' to seperate a Table and Feild in Where statment ie. [Temp]![A]"
    Case 2019
        DisplayError = "SQL statment invalid, no Feild referenced after table name in Select statment"
    Case 2020
        DisplayError = "SQL statment invalid, too many closing brackets in Where statment ie. ( ))"
    Case 2021
        DisplayError = "SQL statment invalid, too many opening brackets in Where statment ie. (( )"
    Case 2022
        DisplayError = "SQL statment invalid, cannot edit a readonly table"
    Case 2023
        DisplayError = ""
    Case 2024
        DisplayError = ""
    Case 2025
        DisplayError = ""
    Case 2026
        DisplayError = ""
    Case 2027
        DisplayError = ""
    Case 2028
        DisplayError = ""
    Case 2029
        DisplayError = ""
    Case 2030
        DisplayError = ""
    Case 2031
        DisplayError = ""
    Case 2032
        DisplayError = ""
    Case 2033
        DisplayError = ""
    Case 2034
        DisplayError = ""
    Case 2035
        DisplayError = ""
    Case 2036
        DisplayError = ""
    Case 2037
        DisplayError = ""
    Case 2038
        DisplayError = ""
    Case 2039
        DisplayError = ""
    Case 2040
        DisplayError = ""
End Select
Debug.Print ErrorNum & vbTab & DisplayError
End Function
