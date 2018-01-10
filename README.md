# mining
Scripts I use for automating various tasks on my crypto mining rig. They only work for nVidia cards as they use the nvidia-smi utility to interrogate the GPUs. The nVidia-SMI utility needs to be installed, but it should be if you installed the drivers with the default options.

I run these scripts from C:\Scripts, but it should be possible to run them from a location of your choosing, as long as they can create a Logs subfolder and files. 

The scripts are:

1) Monitor-ExcavatorRelaunch.vbs - This watches the excavator.exe mining program in NHM2 and just outputs the process name to screen each time it restarts. If a GPU is overclocked too much, excavator will crash and re-launch quite often. This will show you how much the process restarts. If it does it a lot, lower your over clocks.

2) Monitor-GPUs.vbs - This is a "EXE/Script Advanced" Sensor for PRTG Network Monitor. It needs updating with the UUIDs for each of your GPUs. It needs to be run with a parameter. This is the number of the GPU you want to interrogate, starting at 0. This script now automatically finds the GPU UUIDs for all the cards and stores them in the registry. This eliminates the need for the Get-GUIDs.vbs script.

3) Monitor-GPUTotal.vbs - This checks the GPUs aren't in an error state and restarts the rig if a GPU has a fault.

4) Monitor-NetworkConnection.vbs - This checks the rig can ping it's default gateway. If not, it restarts the rig.

5) Monitor-NiceHash.vbs - This checks that the GPUs are running at 80%+ utilisation if NiceHashMinerLegacy has been open for longer than 3mins. If not, then it kills NiceHashMinerLegacy.exe and the miner apps and re-opens NiceHash. If it fails the first test, it checks again 45 seconds later, just to make sure it hasn't hit a time when NHML is switching algorithm.

6) Monitor-PowerLevels.vbs - This checks the power limit for each card is set to 120W. If not, it resets the undervolt using nVidia-SMI.exe. My cards have a habit of forgetting their undervolt settings and if all cards drew power at their full potential, my PSU would pop.

7) Monitor-Power.vbs - This checks the power draw of all the GPUs. If it's greater than 880W it shuts down the computer. This acts as a failsafe in case script 7 fails to set the power limits for some bizarre reason.

8) Monitor-PRTGProbeService.vbs - The PRTG Network Monitor Probe service on my machine has a habit of not starting up after a reboot. This just checks it's running, and starts it up if it isn't.

9) Generate-TaskSchedulerXML.vbs - This VBS has replaced the XML export. It automatically creates the XML file for the Tash Scheduler and imports it to the task scheduler too. Just reboot your machine after you've run this for all the scripts o activate.

10) Monitor-OverClocks.vbs - This checks the current GPU Memory Overclocks and if they are less than the GPU's max overclocks, re-applies an MSI Afterburner profile. If the script is run without a parameter, it will assume profile 1. If you want to apply a different profile, pass the profile number as a parameter (will only accept 1-5) eg wscript.exe Monitor-Overclocks.vbs 3

11) Initialise-ProwlNotifications - Run this script to add your Prowl API key into registry. Prowl Push Notifications will then be sent by the following scripts: Monitor-NiceHash.vbs, Monitor-GPUTotal.vbs, Monitor-Power.vbs, Monitor-PowerLevels.vbs and Monitor-Overclocks.vbs

12) Generate-TaskSchedulerPreFlightChecksXML.vbs - This script generates an XML file to be imported into the task scheduler. It will run the Monitor-PreFlightChecks.vbs script on start up

13) Monitor-PreFlightChecks.vbs - This file runs once at start up and sents a Prowl notification to inform you your machine has booted up (in case it blue screens and auto restarts)

14) Monitor-AwesomeMiner.vbs - After the NiceHash hack I moved to MiningPoolHub and AwesomeMiner. This script is a copy of Monitor-NiceHash.vbs but changed to work with AwesomeMiner

25) Wrapper-ccminer.vbs - Someone asked for a simple script that laucnhed a miner silently and reopened it if it quits. This is that script.

These scripts are a bit rough at the moment. I might get around to polishing them up a bit more if I get time, or if I come across new errors that need fixing, or find a better way of doing things.

If you want to buy me a beer for my troubles and out of the money you make from having a higer mining uptime, then my BTC address is: 1E8kUsm3qouXdVYvLMjLbw7rXNmN2jZesL
