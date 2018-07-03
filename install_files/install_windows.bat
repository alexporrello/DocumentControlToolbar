@ECHO OFF
ECHO Please make sure you are disconnected from VPN before continuing.
@Pause
bitsadmin.exe /transfer "Install Macros" http://github.com/alexporrello/TWBoilerplateMacros/raw/master/binaries/Normal.dotm C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Templates\Normal.dotm
PAUSE