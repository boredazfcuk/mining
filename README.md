# mining
Scripts I use for automating various tasks on my crypto mining rig. They only work for nVidia cards as they use the nvidia-smi utility to interrogate the GPUs. The nVidia-SMI utility needs to be installed, but it should be if you installed the drivers with the default options.

I run these scripts from C:\Scripts and log files are written to C:\Scripts\Logs. You may be able to place them in different folders, but I haven't accounted for folder paths with space names, so that could cause a problem at the moment.

The scripts are:
1) Get-UUIDs.vbs - This retrieves the UUIDs for nVidia GPUs. These are needed for my PRTG Network Monitor Sensor script.

2) Monitor-ExcavatorRelaunch.vbs - This watches the excavator.exe program from NHM2 and just outputs a notification each time it restarts. If a GPU is overclocked too much, excavator will crash and re-launch quite often. This script should be run from cscript.exe and not WScript.

3) Monitor-GPUs.vbs - This is a "EXE/Script Advanced" Sensor for PRTG Network Monitor. It needs updating with the UUIDs for each of your GPUs.

4) Monitor-GPUsTotal.vbs - This checks the GPUs aren't in an error state and restarts the rig if a GPU has a fault.

5) Monitor-NetworkConnection.vbs - This checks the rig can ping it's default gateway. If not, it restarts the rig.

6) Monitor-NiceHash.vbs - This checks that the GPUs are running at 80%+ utilisation if NiceHashMinerLegacy has been open for longer than 3mins. If not, then it kills NiceHashMinerLegacy.exe and the miner apps and re-opens NiceHash.

7) Monitor-PowerLevels.vbs - This checks the power limit for each card is set to 120W. If not, it resets the undervolt using nVidia-SMI.exe. My cards have a habit of forgetting their undervolt settings and if all cards drew power at their full potential, my PSU would pop.

8) Monitor-Power.vbs - This checks the power draw of all the GPUs. If it's greater than 880W it shuts down the computer. This acts as a failsafe in case script 7 fails to set the power limits for some bizarre reason.

9) Monitor-PRTGProbeService.vbs - The PRTG Network Monitor Probe service on my machine has a habit of not starting up after a reboot. This just checks it's running, and starts it up if it isn't.

10) This is an XML export of the scheduled task that I run at startup and repeat every 1 minute. Import it into your task scheduler and then modify it to remove which scripts you don't want to run.

These scripts are a bit rough at the moment. I might get around to polishing them up a bit more if I get time, or if I come across new errors that need fixing, or find a better way of doing things.

If you want to buy me a beer for my troubles, my BTC address is: 1E8kUsm3qouXdVYvLMjLbw7rXNmN2jZesL
