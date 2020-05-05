# New-LyncMeetingWarmup.ps1

This script creates in your Lync Front-End two Scheduled Tasks that fire each time your IIS App Pools are "recycled". The tasks generate a fake meeting join to ensure the sites are always ready for users. This script automates Drago Totev's process, & he deserves all the credit.

**Current version: v1.5 - 12th November 2019.** 

Complaints regarding a slow meeting join process in Lync aren't at all uncommon, and in many cases are actually as a result of IIS's automatic recycling process, itself intended to improve reliability of the IIS  websites.

<a href="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html" target="_blank">Drago Totev has published a fantastic process</a> by which you can ensure your Lync IIS application pools are always sitting  there idling away, ready for a user to join.

In his post, Drago walks you through how to create a scheduled task on your FE's that waits for the Event that's logged each time IIS recycles the Application Pools. When that trigger is seen, the scheduler fires  a script that sends a dummy meeting join request through to IIS which then "warms them up" ready for a real user to join.

I recently followed his setup to add the Schedule to four Front-Ends, and decided by #2 that it would be great if you could just run a PowerShell script to do it all for you automatically. And now you can.

I've opted for one script to do everything, and in this case it serves dual purposes: it's both the script you run to create the Scheduled Events, and it's also the same script that the Scheduler fires to warm  up the app pools. For bonus points it's signed as well (thank you DigiCert), which will hopefully appeal to those who don't like running unsigned scripts.

It's pretty simple to use.

## Creating the Scheduled Tasks

First off, realise that the script needs to be run from the folder you're going to leave it in. The script captures its path in the process of creating the Scheduled task, and the Action it runs is going to look for it  in that location &ndash; so make sure you don't go moving or deleting it after you run the "create" step!

```powershell
.\New-LyncMeetingWarmup.ps1 -CreateTasks -Verbose
```

<img src="https://user-images.githubusercontent.com/11004787/81053634-bf498c00-8f08-11ea-84eb-de5458ece933.png" alt="" width="600" />

Done.

The two Schedules look like this, with only minor differences in the Triggers and Actions to discern between a recycle of the Internal or External site<br /> 

<img src="https://user-images.githubusercontent.com/11004787/81053699-d7b9a680-8f08-11ea-980c-0b738c48371e.png" alt="" width="600" />

## Query the Tasks
You can also run a query to see if or when the tasks have run:

```powershell
.\New-LyncMeetingWarmup.ps1 -GetScheduledTaskInfo
```

(The -verbose switch doesn't reveal any extra info at this stage)

<img src="https://user-images.githubusercontent.com/11004787/81053770-f6b83880-8f08-11ea-997f-078d92f21357.png" alt="" width="600" />

Here's the triggering event in the Event Viewer:

<img src="https://user-images.githubusercontent.com/11004787/81053809-0768ae80-8f09-11ea-8c48-0f19de1fd62b.png" alt="" width="600" />


### Revision History


#### v1.5: 12th November 2019

- Added test for Server 2019. 
- Added my auto-update code. 

#### v1.4: 3rd March 2018

- Thank you @TrevorAMiller for pointing out MS changed the versioning in Server 2016 between Preview and GA. Updated test. 

#### v1.3: 10 February 2015 

- Whoops: fixed a tiny typo in the updated version test that broke it for Win 8.1 & Server 2012 R2. (Sorry) 

#### v1.2: 7 February 2015<span> 

- My colleague Tristan highlighted that the script fails to generate the tasks if your o/s is Server 2008. I've amended it to work as best it can with 2008. It can't create the tasks - that step you'll still need to do manually, however the tasks you create can still just call this script and it will "warm up" your pools for you. I've added more how-to guidance on the blog post. 
- Neatened the CmdletBinding. Makes for a more accurate "get-help" output & blocks unsupported "-WhatIf" and "-Confirm". 
- Tweaked the .EXAMPLES. 

#### v1.1: 17 January 2015

- Realised v1.0 wouldn't work correctly for EE pools, and that the "meetFqdn" isn't actually required. 
- Changed the ScheduledTask Arguments to an execution policy of "AllSigned". (Was unrestricted, leftover from design) &nbsp;Changed write-host to write-output & added "-NoProfile" to $TaskArg (thanks Pat) 

#### v1.0 : 16th January 2015:

- This is the original release 


### Credits

Drago Totev for the original solution: <a title="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html" href="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html"> http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html</a>

Creating a task in PowerShell: <a href="http://www.verboon.info/2013/12/powershell-creating-scheduled-tasks-with-powershell-version-3/"> http://www.verboon.info/2013/12/powershell-creating-scheduled-tasks-with-powershell-version-3/</a>

Tricky Task creation: <a href="http://stackoverflow.com/questions/20108886/scheduled-task-with-daily-trigger-and-repetition-interval"> http://stackoverflow.com/questions/20108886/scheduled-task-with-daily-trigger-and-repetition-interval</a> <br /> <a href="https://p0w3rsh3ll.wordpress.com/2013/07/05/deprecated-features-of-the-task-scheduler/">https://p0w3rsh3ll.wordpress.com/2013/07/05/deprecated-features-of-the-task-scheduler/</a>

<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/new-lyncmeetingwarmup/](https://greiginsydney.com/new-lyncmeetingwarmup/).
