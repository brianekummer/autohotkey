@ECHO OFF
REM --------------------------------------------------------------------------
REM - My Automations is configured using Windows environment variables. This 
REM - batch file will set the necessary environment variables for the current 
REM - Windows USER.
REM - 
REM - Instructions
REM -   1. Modify this batch file as necessary 
REM -   2. Run this batch file
REM -   3. Logout and log back into Windows, or reboot
REM -   4. Make sure your changes are there (e.g. run the command "SET")
REM -   5. Undo your changes to this file or or delete it- you don't need
REM -      them anymore
REM -
REM - To clear a setting, you can do this:
REM -   SETX ENV_VAR_NAME ""
REM --------------------------------------------------------------------------

SETX AHK_URL_JIRA "https://xxxxxxxxxx.atlassian.net"
SETX AHK_MY_PROJECT_KEYS_JIRA "xxxxxx|xxxxxx|xxxxxx"
SETX AHK_DEFAULT_PROJECT_KEY_JIRA "xxxxxx"
SETX AHK_DEFAULT_RAPID_KEY_JIRA "xxx"
SETX AHK_DEFAULT_SPRINT_JIRA "xxxx"

SETX AHK_URL_SOURCE_CODE "https://xxxxxxxxxx"
SETX AHK_URL_TIMESHEET "https://xxxxxxxxxx.com"
SETX AHK_URL_WIKI "https://xxxxxxxxxx"
SETX AHK_URL_PERSONAL_CLOUD "https://xxxxxxxxx"

SETX AHK_MP3_NOISE "x:\xxxxx\xxxxxxx.mp3"

SETX AHK_VIRTUAL_DESKTOP_WALLPAPER_1 "Main    |C:\Users\xxxxxxxxxx\Pictures\wallpaper_main.jpg"
SETX AHK_VIRTUAL_DESKTOP_WALLPAPER_2 "Personal|C:\Users\xxxxxxxxxx\Pictures\wallpaper_personal.jpg"
SETX AHK_VIRTUAL_DESKTOP_WALLPAPER_3 "Temp-1  |C:\Users\xxxxxxxxxx\Pictures\wallpaper_temp_1.jpg"

SETX JOAOAPPS_JOIN_DEVICE_ID="xxxxxxxxxx"
SETX JOAOAPPS_JOIN_API_KEY="xxxxxxxxxx"