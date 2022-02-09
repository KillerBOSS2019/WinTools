![pngwing com_11](https://user-images.githubusercontent.com/55416314/153131886-42c0448d-81f8-49f3-bc38-693ae2341aaf.png)
## WinTool Plugin
- [WinTool Plugin](#wintool-plugin)
    - [Description](#description)
    - [Actions](#actions)
    - [States](#states)
- [Installation Guide](#installation-guide)
- [Settings Overview](#settings-Overview)
- [info](#info)

## Description
WinTool is multi Tools that is made to work with Windows machine.

## Actions
- ![](https://img.shields.io/static/v1?style=for-the-badge&message=VolumeMixer&color=darkgreen&label=Action) [VolumeMixer] Mute/Unmute process volume <br>
- [Action] [VolumeMixer] <h5>Increase/Decrease process volume (slider control available)</h5>
- [Action] [Sound] Change Input/Output Device
- [Action] [UTILITY] Text To Speech: Speak X
- [Action] [UTILITY] Text To Speech: Stop
- [Action] [Mouse] Toggle Mouse Down/Up
- [Action] [Mouse] Stop X, Y Update (Togglable)
- [Action] [Mouse] Teleport Mouse X, Y with delay, animation
- [Action] [Mouse] Advanced Mouse Click
- [Action] [Capture] Display to File / Clipboard 
- [Action] [Capture] Window to File / Clipboard 
- [Action] [Capture] Window to File / Clipboard (*)
- [Action] [Capture] Current Active Window File / Clipboard
- [Action] [UTILITY] Move Application To Monitor Left/Right
- [Action] [UTILITY] Shutdown Computer in (x)
- [Action] [UTILITY] Change Powerplan
- [Action] [UTILITY] System Settings
- [Action] [UTILITY] Change Primary Monitor
- [Action] [UTILITY] Rotate Display / Monitor
- [Action] [UTILITY] Windows Notifcation
- [Action] [UTILITY] Magnifier Glass
- [Action] [UTILITY] On-Screen Keyboard
- [Action] [UTILITY] Emoji Panel
- [Action] [UTILITY] Text / Value to Clipboard
- [Action] [UTILITY] Virtual Desktop: Next / previous
- [Action] [UTILITY] Virtual Desktop: Move Current Window to #

## States
- [State] [Advance Mouse] Current Mouse X pos
- [State] [Advance Mouse] Current Mouse Y pos
- [State] [Sound] Get Current Output Device
- [State] [Sound] Get Current Input Device
- [State] [Application] Get current focused APP
- [State] [Network] Total Bytes Received
- [State] [Network] Total Bytes Send
- [State] [Disk] Total Read
- [State] [Disk] Total Write
- [State] [Window] Active Window Count
- [State] [Power] Active Power Plan
- [State] [System] Computer Uptime 
- [State] [IP] IP Address
- [State] [IP] City
- [State] [IP] Country
- [State] [IP] Region
- [State] [IP] Timezone
- [State] [IP] Internet Provider
- [State] [VolumeMixer] Application Volume (Create states base on Volume Mixer)
- [State] [VolumeMixer] Application is Muted (Create states base on Volume Mixer)

## Installation Guide
1. Go to [Releases](https://github.com/KillerBOSS2019/WinTools/releases) and Latest build should be on the top
2. Then scroll down little until you see "Assets" [Image/Assets.png] Click arrow button
3. And then you should see a file atteched that says `WinTool.V1.0.tpp` something like that. Download it
4. Then open TouchPortal if TouchPortal isnt open. Click on the settings button near the email icon and Select `Import Plug-in...`
5. Then Select the .TPP file that you've downloaded. It should be located in Downloads folder.
6. Next TouchPortal will warn you that you've installed a Plugin called `Windows-Tools`. You can click `Trust Always` if you dont want this to popup everytime you boot. Otherwise click OK for Trust this .
7. Next TouchPortal will popup a Window says `Plug-in Imported successful` You can click `OK`
8. If this was your first Plugin you will need to reboot TouchPortal. Otherwise if plugin doesnt work you could also try reboot TouchPortal
9. That should be all the steps you need to Install the Plugin. If you have issues feel free contect on Discord or Github on issue tab

## Settings Overview

### Windows-Tools Settings
- Update Interval: Network Up-Down(INCOMPLETE)
    - In seconds of how offen that it will update Network data
- Update Interval: Hard Drive
    - In seconds of how offen that it will update infos about storage drive
- Update Interval: Active Monitors
    - In seconds of how offen that it will ipdate active monitors. This is also used for Monitor drop-down list for action
- Update Inerval: Active Windows
    - In seconds of how offen that it will ipdate active window state
- Update Interval: Active Virtual Desktops
    - In seconds of how offen that it will check for Virtual Desktop states

### To access Plugin Settings
1. Click Settings button near the email icon
2. Select `Settings...`
3. It should open Settings window and on the left side you should see `Plug-Ins` tab. Click on that
4. Then you should see a drop-down Select `Windows-Tools`. and you should see all the settings

## Info
Feel free to suggest a pull request for new features, improvements, or documentation. If you are not sure how to proceed with something, please start an Issue at the GitHub repository.
