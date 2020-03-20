'Author: Tahmina Jahan
'Date: 03/12/2020
'Purpose: Check login functionality 
'Test Name: Flight reservation
'Script location:
'Result location: 
'
'*************************************************************************************************************************************
'To find the value that is not declared in the dim 
Option explicit 
'to handle any unexpected error during runtime
On error resume next 
'variable declaration 
Dim n,i,er,er1,expvalmsg

'add new sheet in uft datatable
Datatable.AddSheet ("Flightlogin")
'Select sheet in datatable
Datatable.GetSheet("Flightlogin")
'import excel sheet into the Flightlogin sheet in the datatable
Datatable.ImportSheet "C:\Users\jahan\Desktop\Automation advance\flightreservation\flightlogindata.xlsx",1,"Flightlogin"
'Count number of avaiable rows in testing sheet
n=Datatable.GetSheet("Flightlogin").GetRowCount 
For i = 1 To n

Call fnopenapp()

Dialog("Login").WinEdit("Agent Name:").Set Datatable ("Agent_Name","Flightlogin") @@ hightlight id_;_4001062_;_script infofile_;_ZIP::ssf2.xml_;_
Dialog("Login").WinEdit("Password:").Set Datatable ("Password","Flightlogin") @@ hightlight id_;_2952436_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("Login").WinButton("OK").Click @@ hightlight id_;_3539656_;_script infofile_;_ZIP::ssf4.xml_;_

If Dialog("Login").Dialog("Flight Reservations").Exist  Then
Dialog("Login").Dialog("Flight Reservations").Check CheckPoint("Flight Reservations") @@ hightlight id_;_1903514_;_script infofile_;_ZIP::ssf6.xml_;_
er=Dialog("Login").Dialog("Flight Reservations").Static("Agent name must be at").GetROProperty("text") @@ hightlight id_;_3148982_;_script infofile_;_ZIP::ssf7.xml_;_
datatable.Value("Actual_validation_message","Flightlogin")= er
Dialog("Login").Dialog("Flight Reservations").WinButton("OK").Click
Dialog("Login").WinButton("Cancel").Click @@ hightlight id_;_1378372_;_script infofile_;_ZIP::ssf12.xml_;_

Else
Window("Flight Reservation").Check CheckPoint("Flight Reservation") @@ hightlight id_;_3804388_;_script infofile_;_ZIP::ssf10.xml_;_
er1 = Window("Flight Reservation").GetROProperty("text") @@ hightlight id_;_3804388_;_script infofile_;_ZIP::ssf11.xml_;_
datatable.Value("Actual_validation_message",dtglobalsheet)= er1

Window("Flight Reservation").Check CheckPoint("Flight Reservation_3") @@ hightlight id_;_2034308_;_script infofile_;_ZIP::ssf13.xml_;_
Window("Flight Reservation").WinObject("Flight Schedule:").Check CheckPoint("Flight Schedule:") @@ hightlight id_;_3016418_;_script infofile_;_ZIP::ssf14.xml_;_
Window("Flight Reservation").Close
End If
'Get the expected result from uft datatable and keep in a variable
expvalmsg=Datatable("Expected_validation_message","Flightlogin")
'Compare actual result and expected result and varify the status
If instr(expvalmsg,er) Then
	Datatable("Status","Flightlogin")= "Unsuccessfull login with invalid input" &" "& "Error message displayed and captured" &" "& "pass"  
ElseIf instr(expvalmsg,er1) Then
	Datatable("Status","Flightlogin")="Successfull login with valid input" &" "& "No error message displayed" &" "& "pass" 
Else
	Datatable("Status","Flightlogin")="Not matching with expected result"&" "&"fail"

End if 
datatable.SetNextRow 
Next

'Exporting datatable from uft datatable to external excel sheet
Datatable.ExportSheet "C:\Users\jahan\Desktop\Automation advance\flightreservation\flightlogindata.xlsx","Flightlogin"






















