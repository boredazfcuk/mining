# mining
Scripts I use for automating various tasks on my crypto mining rig. They only work for nVidia cards as they use the nvidia-smi utility to interrogate the GPUs. The nVidia-SMI utility needs to be installed, but it should be if you installed the drivers with the default options.

I run these scripts from C:\Scripts, but it should be possible to run them from a location of your choosing, as long as they can create a Logs subfolder and files. 

The scripts are:
1) Get-UUIDs.vbs - This retrieves the UUIDs for nVidia GPUs. These are needed for my PRTG Network Monitor Sensor script.

2) Monitor-ExcavatorRelaunch.vbs - This watches the excavator.exe mining program in NHM2 and just outputs the process name to screen each time it restarts. If a GPU is overclocked too much, excavator will crash and re-launch quite often. This will show you how much the process restarts. If it does it a lot, lower your over clocks.

3) Monitor-GPUs.vbs - This is a "EXE/Script Advanced" Sensor for PRTG Network Monitor. It needs updating with the UUIDs for each of your GPUs. It needs to be run with a paramete. This is the number of the GPU you want to interrogate, starting at 0.

4) Monitor-GPUsTotal.vbs - This checks the GPUs aren't in an error state and restarts the rig if a GPU has a fault.

5) Monitor-NetworkConnection.vbs - This checks the rig can ping it's default gateway. If not, it restarts the rig.

6) Monitor-NiceHash.vbs - This checks that the GPUs are running at 80%+ utilisation if NiceHashMinerLegacy has been open for longer than 3mins. If not, then it kills NiceHashMinerLegacy.exe and the miner apps and re-opens NiceHash. If it fails the first test, it checks again 45 seconds later, just to make sure it hasn't hit a time when NHML is switching algorithm.

7) Monitor-PowerLevels.vbs - This checks the power limit for each card is set to 120W. If not, it resets the undervolt using nVidia-SMI.exe. My cards have a habit of forgetting their undervolt settings and if all cards drew power at their full potential, my PSU would pop.

8) Monitor-Power.vbs - This checks the power draw of all the GPUs. If it's greater than 880W it shuts down the computer. This acts as a failsafe in case script 7 fails to set the power limits for some bizarre reason.

9) Monitor-PRTGProbeService.vbs - The PRTG Network Monitor Probe service on my machine has a habit of not starting up after a reboot. This just checks it's running, and starts it up if it isn't.

10) This is an XML export of the scheduled task that I run at startup and repeat every 1 minute. Import it into your task scheduler and then modify it to remove which scripts you don't want to run.

These scripts are a bit rough at the moment. I might get around to polishing them up a bit more if I get time, or if I come across new errors that need fixing, or find a better way of doing things.

If you want to buy me a beer for my troubles and out of the money you make from having a higer mining uptime, then my BTC address is: 1E8kUsm3qouXdVYvLMjLbw7rXNmN2jZesL
