strComputer = "."
Set objWMI1 = GetObject("winmgmts://" & strComputer & "/root/wmi")
Set temp = objWMI1.instancesof("MSAcpi_ThermalZoneTemperature")
msgbox temp