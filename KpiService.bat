mkdir C:\KPI
cd C:\KPI
git pull
python C:\Kpi\KpiCollector.py
mkdir c:\inetpub
mkdir c:\inetpub\wwwroot
copy RotatingDisplay.html c:\inetpub\wwwroot\.
copy KPI.html c:\inetpub\wwwroot\.
copy SexKPIer.html c:\inetpub\wwwroot\.
copy KPI-URLs.csv c:\inetpub\wwwroot\.
copy *.js c:\inetpub\wwwroot\.


