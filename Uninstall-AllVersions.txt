##there are two types of MSO uninstall
## (1) uninstall office proplus, project and visio (x64, x86)
"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall ProPlus /config "\\10.10.10.10\New\uninstall2016.xml" & 
"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall PRJPRO /config "\\10.10.10.10\New\uninstall2016-Project-.xml" &  
"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall VISPRO /config "\\10.10.10.10\New\uninstall2016-Visio-.xml" & 
"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall VISPRO /config "\\10.10.10.10\New\uninstall2016-Visio-.xml" & 
"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall ProPlus /config "\\10.10.10.10\New\uninstall2016.xml" &
"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe" /uninstall PRJPRO /config "\\10.10.10.10\New\uninstall2016-Project-.xml" &

## (2) uninstall office proplus, project and visio (x64, x86)
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=ProplusRetail.16_en-us_x-none culture=en-us version.16=16.0 &
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=VisioProRetail.16_en-us_x-none culture=en-us version.16=16.0 &
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=ProjectProRetail.16_en-us_x-none culture=en-us version.16=16.0 &
"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=ProplusRetail.16_en-us_x-none culture=en-us version.16=16.0 &
"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=ProjectProRetail.16_en-us_x-none culture=en-us version.16=16.0 & 
"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=VisioProRetail.16_en-us_x-none culture=en-us version.16=16.0 &

##unistall proffing toolds
MsiExec.exe /qn /norestart /X{90160000-001F-0401-1000-0000000FF1CE} &
MsiExec.exe /qn /norestart /X{90150000-001F-0401-0000-0000000FF1CE} &

