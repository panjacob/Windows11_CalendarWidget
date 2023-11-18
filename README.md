# Windows11_CalendarWidget
Minimal calendar app displays schedule from Google Calendar API.

## 1. Motivation
I couldn't find a good enough solution to display the schedule of classes I planned in Google Calendar. I wanted to create a widget that is similar, to those available on smartphones.

## 2. How it works?

App displays 10 recent activities from calendar. <br/>
<img src="https://download.panjacob.pl/1Big.png" alt="drawing" width="400"/>

<br/>
You can click calendar icon on the tray to open settings or close application.
<br/>

![Zrzut ekranu 2023-11-13 170053](https://download.panjacob.pl/2Tray.png)

<br/>

Right now only a few options are customizable. You can change: dimensions and position of a window. Additionaly you can logout, update events and disable/enable startup.<br/>
![Zrzut ekranu 2023-11-18 163022](https://download.panjacob.pl/3Settings.png)
<br/>

## 3. How to run?
### 3.1. Download current release <br/> 
![Zrzut ekranu 2023-11-18 163022](https://download.panjacob.pl/4Releases.png)
### 3.2. Unpack and run CalendarWidget.exe
### 3.3. Login to your Google account. If there is no future events then widget is not visible (to be fixed). 
### 3.3 Tips:
- You can move widget by holding click on a widget and moving mouse. 
- It's possible to enable startup in settings. If file has been moved it won't work until startup is disabled and enabled again.
- Change width and size of widget depending of your screen resolution.

## 4. How to compile?
- For this project python 3.10 was used.
- `pip install requirements.txt`
- create file **credentials.py**. It should look like this: <br/>
![credentials.py](https://download.panjacob.pl/5Credentials.png)
- To be continued...