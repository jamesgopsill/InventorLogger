# InventorLogger
[Dr. James Gopsill](http://jamesgopsill.github.io)

<img src="sequence-transitions.png" width="300px">

This GitHub Repository contains the VB code for an add-in to AutoDesk Inventor that enables researchers to log user interaction with the software. The tool was built as part of a study to understand how engineers approach the design of product models within a Computer Aided Design system and the results were reported at the International Conference of DESIGN 2016.

The research was funded by the EPSRC through the [Language of Collaborative Manufacturing](http://locm.blogs.ilrt.org) project.


The analysis of user interactions within Computer Aided Design tools is a fascinating research topic and is one we want to actively encourage the community to get involved in. Please feel free to clone, download and/or fork this software for your own studies. All we ask is that you cite our work so that we can continue to develop tools for the community.

Also, please let us know if you have used the logger and/or derivation of the logger in your studies! We would like to keep and up-to-date list of publications that use these tools so future researchers can read up and how they have been used.

# Citation

```latex

@inproceedings{gopsill2016,
  title={Computer Aided Design User Interaction as a Sensor for Monitoring Engineers and the Engineering Design Process},
  author={Gopsill, J.A. and Snider, C. and Shi, L. and Hicks, B.J.},
  booktitle={International Conference on DESIGN},
  year={2016},
  url={https://jamesgopsill.github.io/Publications/papers/conference/design2016/design2016.pdf}
}

```

# Installing InventorLogger

The Inventor Logger has already been compiled and should work on most Windows systems with AutoDesk Inventor. Once you have downloaded the folder there are only three steps required to get it working.

## Step 1: Copy the .dll

You need to find and copy the `\InventorLogger1\bin\Debug\InventorLogger1.dll` file into the AutoDesk Inventor bin folder. This is usually loacted at your install path for Inventer `<InstallPath>\bin\`.

## Step 2: Copy the .addin file

Now, to get Inventor to look for and load the .dll at start-up, we need to find and copy the .add-in file. This is located at `\InventorLogger1\AutoDesk.InventorLogger1.Inventor.addin`. This then needs to be copied to the Inventor add-in folder and is usually loacted at `C:\ProgramData\AutoDesk\Inventor Add-ins\`.

## Step 3: Create InventorLogs folder

When you load Inventor, the InventorLogger will start-up and look for a folder called `C:\ProgramData\InventorLogs\`. You will need to create this folder so that it create a log file everytime Inventor is loaded.

# Logging with the InventorLogger

Everytime Inventor is opened, the InventorLogger should automatically create a new log file in `C:\ProgramData\InventorLogs\`. To check if the logger is working, you can click on the add-ins menu in AutoDesk Inventor to see if it has been loaded.

# Disabling the InventorLogger

To disable the InventorLogger, all you need to do is go into the add-ins menu within AutoDesk Inventor and deselect the InventorLogger. You can also disable to InventorLogger from loading automatically at start-up.

# Removing the InventorLogger

Removing the InventorLogger is as simple as deleting the `.dll` and `.addin` files from your PC.

<!--
# The AutoDesk API
-->

# List of Publications

[Gopsill, J.A., Snider, C.M., Shi, L. & Hicks, B.J. (2016). Computer Aided Design User Interaction as a Sensor for Monitoring Engineers and the Engineering Design Process}. In: *Proceedings of the International Conference on DESIGN*](https://jamesgopsill.github.io/Publications/papers/conference/design2016/design2016.pdf)
