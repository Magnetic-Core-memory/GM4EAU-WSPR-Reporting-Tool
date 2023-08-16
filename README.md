
WATT - Whisper Analysis and Traffic Tool     
Reviewed/updated Aug2023.

WATT is an Excel spreadsheet based tool (developed using Microsoft VBA.) It’s built to view and plot WSPR data held by wsprnet.org. Traffic can be animated or re-played over a time-line up to 30 days. Full Excel spreadsheet functionality is also available to allow ad-hoc analysis of the queried data, and to enhance the tool with new functionality to meet specific requirements. All filtered/selected data can be dispayed and animated on a world-map. 

Pre-requisites: Microsoft Excel 2010 or higher. All 32 and 64 Bit versions and Office 365  versions supported at Aug 2023. Not tested on Windows11.

WATT website:  www.gm4eau.com

The tool offers WSPR data queries using a series of on-screen buttons and forms and by using standard spreadsheet filtering functionality. Additionally, traffic data can be plotted on a world map to view activity over a period of time (e.g. the past 24 hours, past month, etc.) As the timeline progresses, the traffic within each period is plotted (rather like a Weather Rain Radar map.) Users can add their own analysis using standard Excel functionality to provide specific charts/graphs or other results. Full documentation and examples are included within the tool.

The VBA source code is not password protected within the spreadsheet. This are  approximately 6000 lines of VBA code, and this is essentially bug free after 3 years of usage by the Amateur Radio community. The code can fail in some extreme user induced conditions since defensive coding within Excel VBA is difficult and costly in effort to develop, however for normal usage the software performs very well: Button operations can in theory operate concurrently. Some Excel bugs cannot be accommodated and workarounds have been applied.

![image](https://user-images.githubusercontent.com/41966359/147089224-f43e49e4-6566-4e4b-a16a-618ec743978a.png)
![image](https://user-images.githubusercontent.com/41966359/147089417-6ec2d8e5-d7c4-4864-a5eb-fa1bb72f510e.png)
![image](https://user-images.githubusercontent.com/41966359/147089459-480e9f03-66b9-46d9-bf92-9f7881141629.png)
![image](https://user-images.githubusercontent.com/41966359/147089531-de653014-403d-47b1-bb2c-cb6d0de09f52.png)
![image](https://user-images.githubusercontent.com/41966359/147089543-7c10560c-4106-4e1f-a187-0f238029f4d5.png)
![image](https://user-images.githubusercontent.com/41966359/147089574-b425ce64-0544-4b81-b7d8-dec134a9aa2d.png)
![image](https://user-images.githubusercontent.com/41966359/147089583-41a7ee85-e1e7-41ba-80a3-c9c6370d914f.png)
![image](https://user-images.githubusercontent.com/41966359/147089601-92f5da8b-7030-47de-a772-b5c5a1a59d40.png)
![image](https://user-images.githubusercontent.com/41966359/147089618-fff98849-df8c-4f5c-ad57-a003a03beaa6.png)
![image](https://user-images.githubusercontent.com/41966359/147089632-1608b2a1-249c-44aa-9078-e63691ac75dc.png)

WATT currently constrains the lines of query data to avoid loading the wsprnet.org database, however the software is designed to handle higher volumes and this can be adjusted. The unique feature of WATT ( in addition to the more routine plotting of traffic) is the Timeline which filters data into time slots and animates each in turn on a world map – all data on the world map is clickable for further information.

Software Compatibility:

To operate this Microsoft Excel spreadsheet software, you will need to authorise Data links and Excel macros to run.
The software has been tested on Microsoft Excel 2010, 2013, 2016, and Office365 (at Jul 19 with Excel 2019.) on Windows 7 and Windows 10 up to August 2023 builds. Not tested with Windows 11.
32 Bit and 64 bit Excel  are supported.  
Open Office and similar are not supported (they don’t provide VBA).  Apple MAC is not supported.

If the software does fail, simply re-open the spreadsheet. The code has not been designed to be highly robust given the complexity of implementing this logic within VBA, however the key features are protected from User error. All features work correctly for ‘normal’ use.

Operating with 1000-2000 lines of WSPR data should not incur any processing delays of more than a few seconds to request data or draw a Map. The timeline may take much longer if a very high number of time slots are requested. The software will run up to 50,000+  lines (for archived data  analysis activities) however the code is currently software  limited at 10,000 lines.

Author contact details are available within the About button on the DATA tab within the WATT tool.
