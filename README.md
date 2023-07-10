
# Windows-Tools
- [Windows Tools](#Windows-Tools)
  - [Description](#description)
  - [Features](#Features)
    - [Actions](#actions)
        - [Python Examples](#com.github.KillerBOSS2019.TouchPortal.plugin.WinTool.mainaction)
        - [Advanced Mouse](#com.github.KillerBOSS2019.TouchPortal.plugin.WinTool.AdvancedMouseaction)
    - [Connectors](#connectors)
  - [Bugs and Support](#bugs-and-suggestion)
  - [License](#license)
  
# Description

This documentation generated for Windows Tools V300 with [Python TouchPortal SDK](https://github.com/KillerBOSS2019/TouchPortal-API).
# Features

## Actions
<details open id='com.github.KillerBOSS2019.TouchPortal.plugin.WinTool.mainaction'><summary><b>Category:</b> Python Examples <ins>(Click to expand)</ins></summary><table>
<tr valign='buttom'><th>Action Name</th><th>Description</th><th>Format</th><th nowrap>Data<br/><div align=left><sub>choices/default (in bold)</th><th>On<br/>Hold</sub></div></th></tr>
<tr valign='top'><td>Advanced Program launcher</td><td> </td><td>Launch[1][2]</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Pick One</b> Possible choices: ['Steam', 'Microsoft', 'Other']</li>
<li>Type: choice &nbsp; 
Default: <b>Rocket League</b></li>
</ol></td>
<td align=center>No</td>
<tr valign='top'><td>Change Powerplan</td><td> </td><td>Change current powerplan to[1]</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Balanced</b></li>
</ol></td>
<td align=center>No</td>
</tr></table></details>
<details  id='com.github.KillerBOSS2019.TouchPortal.plugin.WinTool.AdvancedMouseaction'><summary><b>Category:</b> Advanced Mouse <ins>(Click to expand)</ins></summary><table>
<tr valign='buttom'><th>Action Name</th><th>Description</th><th>Format</th><th nowrap>Data<br/><div align=left><sub>choices/default (in bold)</th><th>On<br/>Hold</sub></div></th></tr>
<tr valign='top'><td>Mouse Hold/Release mouse button</td><td> </td><td>[1]mouse button[2]</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Hold</b> Possible choices: ['Hold', 'Release']</li>
<li>Type: choice &nbsp; 
Default: <b></b> Possible choices: ['Left', 'Middle', 'Right']</li>
</ol></td>
<td align=center>No</td>
<tr valign='top'><td>Mouse Click</td><td> </td><td>[1]Click[2]Times with interval[3]</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Left</b> Possible choices: ['Right', 'Middle', 'Left']</li>
<li>Type: number &nbsp; 
Default: <b>3</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: number &nbsp; 
Default: <b>1</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647 &nbsp; <b>Allow Decimals:</b> True</li>
</ol></td>
<td align=center>No</td>
<tr valign='top'><td>Mouse scroll</td><td> </td><td>Scroll Mouse[1]by[2]ticks</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Hold</b> Possible choices: ['UP', 'DOWN', 'LEFT', 'RIGHT']</li>
<li>Type: number &nbsp; 
Default: <b>120</b> &nbsp; <b>Min Value:</b> 1 &nbsp; <b>Max Value:</b> 2147483647</li>
</ol></td>
<td align=center>No</td>
<tr valign='top'><td>Move mouse</td><td> </td><td>Mouse [1]X[2]Y[3]in duration[4]</td><td><details><summary><ins>Click to expand</ins></summary><ol start=1>
<li>Type: choice &nbsp; 
Default: <b>Hold</b> Possible choices: ['To', 'Move']</li>
<li>Type: number &nbsp; 
Default: <b>10</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: number &nbsp; 
Default: <b>10</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: number &nbsp; 
Default: <b>2</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
</ol></td>
</details><td align=center>No</td>
<tr valign='top'><td>Drag Mouse</td><td> </td><td>[1][1]X[2]Y[3]in duration[4]with button [5]</td><td><details><summary><ins>Click to expand</ins></summary><ol start=1>
<li>Type: choice &nbsp; 
Default: <b>Hold</b> Possible choices: ['DragTo', 'Drag']</li>
<li>Type: number &nbsp; 
Default: <b>10</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: number &nbsp; 
Default: <b>10</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: number &nbsp; 
Default: <b>2</b> &nbsp; <b>Min Value:</b> -2147483648 &nbsp; <b>Max Value:</b> 2147483647</li>
<li>Type: choice &nbsp; 
Default: <b>Left</b> Possible choices: ['Left', 'Middle', 'Right']</li>
</ol></td>
</details><td align=center>No</td>
</tr></table></details>
<br>

## Connectors
<details open><summary><b>Category:</b> Python Examples <ins>(Click to expand)</ins></summary><table>
<tr valign='buttom'><th>Slider Name</th><th>Description</th><th>Format</th><th nowrap>Data<br/><div align=left><sub>choices/default (in bold)</th></tr>
<tr valign='top'><td>Slider Mouse scroll (HScroll/Scroll)</td><td> </td><td>[1]Scroll mouse at scroll speed [2]x and Reverse[3]</td><td><ol start=1><li>Type: choice &nbsp; 
Default: <b>Forward/Backward</b> Possible choices: ['Forward/Backward', 'Up/Down']</li>
<li>Type: number &nbsp; 
Default: <b>20</b> &nbsp; <b>Min Value:</b> 1 &nbsp; <b>Max Value:</b> 120</li>
<li>Type: choice &nbsp; 
Default: <b>False</b> Possible choices: ['True', 'False']</li>
</ol></td>
</table></details>
<br>

# Bugs and Suggestion
Open an issue on github or join offical [TouchPortal Discord](https://discord.gg/MgxQb8r) for support.


# License
This plugin is licensed under the [GPL 3.0 License] - see the [LICENSE](LICENSE) file for more information.

