'Login Scriptt
systemutil.Run "C:\Users\tareq.ahmed\Desktop\Documents\Network\Recovery Tool - 2019\UFT Flight App\flight4b.exe"
Dialog("Login").WinEdit("Agent Name:").Set "mercury" @@ hightlight id_;_199578_;_script infofile_;_ZIP::ssf1.xml_;_
Dialog("Login").WinEdit("Password:").Set "mercury" @@ hightlight id_;_658208_;_script infofile_;_ZIP::ssf2.xml_;_
Dialog("Login").WinButton("OK").Click @@ hightlight id_;_723770_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Flight Reservation").Close