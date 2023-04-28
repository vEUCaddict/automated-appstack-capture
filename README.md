# FULLY AUTOMATED VMWARE APP VOLUMES APPSTACK CREATION

A couple of years ago, I was running a project where I needed to appstack 150 plus applications. In that time VMware had not an automated way of creating the appstacks, so I was starting to analyze the process and created a PowerShell script for it. 
In the recent versions VMware created some automation though, but I still received so now and then on the socials a request to share my script, that they watched my YouTube video from the VMware Code Connect session 2022 (https://youtu.be/NbERAIkky-s).<br />
I would like to give some credits to community member Chris Twiest (he started with a basic script based on App Volumes 2.x) and my former employer ITQ for stimulating and give the time to create this.
<br />
<br />
## Before you begin with the script, there are some prerequisites!!!
Follow the detailed procedure for the prerequisites consult my blog (https://veucaddict.com).<br />
There I will explain why certain things needs to be done or set.<br />

Below is a enumeration for the steps:
<br />
- To talk to VMware vCenter and ESXi via PowerShell you need to have the VMware PowerCLI installed.
- An Active Directory service account, with the appropriate permissions.
- DFS/NTFS share to store the software and the scripts, your App Volumes Capture VM needs to have access to it  via the Active Directory service account.
- The Active Directory service account needs permissions set in the App Volumes Manager.
- Create a custom VMware vCenter role with the Active Directory service account associated to the role and set the role permissions to a specific vSphere Folder.
- A best practice is to disable Anti-virus and Anti-malware software, disable the Windows Firewall and disable Microsoft/Windows Updates on the App Volumes Capture VM.
- The App Volumes Capture VM must be domain joined, place the computer object in a specific OU with block group policy inheritance.
- By default a Windows computer doesn't logon automatically anymore (unless you force too).
<br />
- Set some specific Active Directory Group Policies for: 
	- Maximum machine account password age;
	- Windows Remote Management service (WinRM) Startup Mode; 
	- Windows Firewall rules for WinRM (just to be sure);
	- Allow delegating fresh credentials;
	- Allow remote server management through WinRM;
	- Allow unencrypted traffic;
	- Disallow WinRM from storing RunAs credentials;
	- Windows AutoLogon registry keys;
	- Add packaging/capturing users to the built-in local administrators group.
<br />
<br />
### Tips & Remarks
- VMware doesn't support AppStacks created this way... So if you have any findings, please try creating an AppStack manually before calling support, that a certain application doesn't work.
- A friendly note, this setup works out for me, tested it with a lot of freeware applications in non-nested virtualization environment. The VM which I use for creating AppStacks was running on a Intel NUC 10 with ESXi 7.x on it. 
- In the script there is a reboot time of the Capture VM specified, maybe you need to finetune this reboot time for your own environment. My advice is, don't decrease the number only increase.
- For any questions or recommendations, please leave a message on GitHub/Blog/LinkedIn/Twitter.
<br />
<br />
### Commands to run the script
Don't forget to edit the variables in the variable section to your environment needs.<br />
Open a PowerShell window (not preferred running as administrator).<br />
<br />
Template:<br />

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
cd %scriptlocation%
.\FullyAutomated-AppStack-Creation-%APP%-%VERSION%.ps1 -AV_CAP_VM_NAME %VMNAME%
```

Example:<br />
```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
cd D:\Sources\AppStack
.\FullyAutomated-AppStack-Creation-PuTTY-0.78.ps1 -AV_CAP_VM_NAME "W11X64AVCAP99"
```

Reach out to me on the following social channels:<br />
GitHub:   https://github.com/veucaddict<br />
Blogsite: https://veucaddict.com<br />
LinkedIn: https://www.linkedin.com/in/sidneylaan<br />
Twitter:  https://twitter.com/vcloudaddict<br />