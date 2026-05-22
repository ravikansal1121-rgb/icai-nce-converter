@echo off
echo  Setting up firewall rule for ICAI NCE Converter (port 5001)...
echo  (Requires Administrator privileges)
echo.
netsh advfirewall firewall add rule name="ICAI NCE Converter" dir=in action=allow protocol=tcp localport=5001
echo.
echo  Done! Other devices on your WiFi can now access the tool.
pause
