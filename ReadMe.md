## Synopsis

The purpose of this PowerShell scripting tool is to improve the procedures that are performed whenever an employee 
leaves the company. There are many task that HR and IT need to perform and often they are manual and tedious. Also 
it is sometimes not clear which are completed and which are pending. This set of scripts aims to greatly improve 
that situation for all parties involved.

Note, this script is setup to be run from another script called the "Launch Pad" script. 

Also, it relies on having 
a system on premises that runs a scheduled task that actually processes the requests that this script creates.

Also, it relies on having a SharePoint Online instance setup, which the server-side script will reference.


## Code Example

One way to bring up this tool to avoid having to use the command line is to create a Windows shortcut to PowerShell.exe 
and then call the LaunchPad.ps1 file...

C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe -File "\\server\share\scripts\PowerShell-LaunchPad.ps1"

*If you need to bypass a local workstation's execution policy, you could include the following in the shortcut target...
-ExecutionPolicy Bypass 

The LaunchPad.ps1 script looks in the $global:lpLaunchFiles directory and pulls up any .ps1 files to populate the drop-down 
with choices.

## Motivation

The IT and HR team would always have feedback and confusion in areas of employee termination or "offboarding". There were many tasks 
performed that were repetitive and time consuming, and often time reminders were needed for tasks that are deferred x # of days out.

## Installation

Prior to using this script, you need to...
1) have a SharePoint Online instance setup, which you have access to administer.
2) have a scheduled task created on a server, with the server counterparts to this script.
3) the service account that runs the above mentioned script needs to also be a SharePoint Online user, with admin rights in the workspace being used.
4) this script refers to a network share, where the XML files will be saved to.

## API Reference

No API here. <sounds of crickets>

## Tests

No testing info here. <sounds of crickets>

## Contributors

Just a solo script project by moi, Greg Besso. Hi there :-)

## License

Copyright (c) 2016 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.