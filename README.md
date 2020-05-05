# New-LyncMeetingWarmup.ps1

This script creates in your Lync Front-End two Scheduled Tasks that fire each time your IIS App Pools are "recycled". The tasks generate a fake meeting join to ensure the sites are always ready for users. This script automates Drago Totev's process, &amp; he deserves all the credit.

<p>&nbsp;</p>
<p><span style="color: #ff0000;"><strong><span style="font-size: small;">Current version: v1.5 - 12th November 2019.</span></strong></span></p>
<p><span style="color: #ff0000;"><strong><span style="font-size: small;"><br /> </span></strong></span></p>
<p><span style="font-size: small;">Complaints regarding a slow meeting join process in Lync aren&rsquo;t at all uncommon, and in many cases are actually as a result of IIS&rsquo;s automatic recycling process, itself intended to improve reliability of the IIS  websites.</span></p>
<p><span style="font-size: small;"><a href="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html" target="_blank">Drago Totev has published a fantastic process</a> by which you can ensure your Lync IIS application pools are always sitting  there idling away, ready for a user to join.</span></p>
<p><span style="font-size: small;">In his post, Drago walks you through how to create a scheduled task on your FE&rsquo;s that waits for the Event that&rsquo;s logged each time IIS recycles the Application Pools. When that trigger is seen, the scheduler fires  a script that sends a dummy meeting join request through to IIS which then &ldquo;warms them up&rdquo; ready for a real user to join.</span></p>
<p><span style="font-size: small;">I recently followed his setup to add the Schedule to four Front-Ends, and decided by #2 that it would be great if you could just run a PowerShell script to do it all for you automatically. And now you can.</span></p>
<p><span style="font-size: small;">I&rsquo;ve opted for one script to do everything, and in this case it serves dual purposes: it&rsquo;s both the script you run to create the Scheduled Events, and it&rsquo;s also the same script that the Scheduler fires to warm  up the app pools. For bonus points it&rsquo;s signed as well (thank you DigiCert), which will hopefully appeal to those who don&rsquo;t like running unsigned scripts.</span></p>
<p><span style="font-size: small;">It&rsquo;s pretty simple to use.</span></p>
<h3>Creating the Scheduled Tasks</h3>
<p><span style="font-size: small;">First off, realise that the script needs to be run from the folder you&rsquo;re going to leave it in. The script captures its path in the process of creating the Scheduled task, and the Action it runs is going to look for it  in that location &ndash; so make sure you don&rsquo;t go moving or deleting it after you run the &ldquo;create&rdquo; step!</span></p>
<p><span style="font-size: small;">.\New-LyncMeetingWarmup.ps1 -CreateTasks -Verbose</span></p>
<pre><span style="font-size: small;"><br /></span></pre>

<img src="https://user-images.githubusercontent.com/11004787/81053634-bf498c00-8f08-11ea-84eb-de5458ece933.png" alt="" width="600" />
<p><span style="font-size: small;">Done.</span></p>
<p><span style="font-size: small;">The two Schedules look like this, with only minor differences in the Triggers and Actions to discern between a recycle of the Internal or External site<br /> </span></p>

<img src="https://user-images.githubusercontent.com/11004787/81053699-d7b9a680-8f08-11ea-980c-0b738c48371e.png" alt="" width="600" />

<h3>Query the Tasks</h3>
<p><span style="font-size: small;">You can also run a query to see if or when the tasks have run:</span></p>
<pre><span style="font-size: small;">.\New-LyncMeetingWarmup.ps1 &ndash;GetScheduledTaskInfo </span></pre>
<p><span style="font-size: small;">(The &ndash;verbose switch doesn&rsquo;t reveal any extra info at this stage)</span></p>

<img src="https://user-images.githubusercontent.com/11004787/81053770-f6b83880-8f08-11ea-997f-078d92f21357.png" alt="" width="600" />
<p><span style="font-size: small;">Here&rsquo;s the triggering event in the Event Viewer:</span></p>

<img src="https://user-images.githubusercontent.com/11004787/81053809-0768ae80-8f09-11ea-8c48-0f19de1fd62b.png" alt="" width="600" />
<p>&nbsp;</p>
<h3><span style="font-size: small;">Revision History</span></h3>
<p><span style="font-size: small;"><br /> </span></p>
<p><span style="font-size: small;">v1.5: 12th November 2019</span></p>
<ul>
<li><span style="font-size: small;">Added test for Server 2019.</span> </li>
<li><span style="font-size: small;">Added my auto-update code.</span> </li>
</ul>
<p>&nbsp;</p>
<p><span style="font-size: small;">v1.4: 3rd March 2018</span></p>
<ul>
<li><span style="font-size: small;">Thank you @TrevorAMiller for pointing out MS changed the versioning in Server 2016 between Preview and GA. Updated test.</span> </li>
</ul>
<p>&nbsp;</p>
<p><span style="font-size: small;">v1.3: 10 February 2015 </span></p>
<ul>
<li><span style="font-size: small;">Whoops: fixed a tiny typo in the updated version test that broke it for Win 8.1 &amp; Server 2012 R2. (Sorry)</span> </li>
</ul>
<p>&nbsp;</p>
<p><span style="font-size: small;">v1.2: 7 February 2015<span> </span></span></p>
<ul>
<li><span style="font-size: small;">My colleague Tristan highlighted that the script fails to generate the tasks if your o/s is Server 2008.</span><span style="font-size: small;"> </span><span style="font-size: small;">I've amended it to work as best it can with 2008. It can't create the tasks - that step you'll</span><span style="font-size: small;"> </span><span style="font-size: small;">still need to do manually, however the tasks you create can still just call this script and it will</span><span style="font-size: small;"> </span><span style="font-size: small;">"warm up" your pools for you. I've added more how-to guidance on the blog post.</span> </li>
<li><span style="font-size: small;">Neatened the CmdletBinding. Makes for a more accurate "get-help" output &amp; blocks unsupported "-WhatIf" and "-Confirm".</span> </li>
<li><span style="font-size: small;">Tweaked the .EXAMPLES.</span> </li>
</ul>
<p><span style="font-size: small;">v1.1: 17 January 2015</span></p>
<ul>
<li><span style="white-space: pre;">&nbsp;</span><span style="font-size: small;">Realised v1.0 wouldn't work correctly for EE pools, and that the "meetFqdn" isn't actually required.</span> </li>
<li><span style="white-space: pre;">&nbsp;</span><span style="font-size: small;">Changed the ScheduledTask Arguments to an execution policy of "AllSigned". (Was unrestricted, leftover from design) </span><span style="white-space: pre;">&nbsp;</span><span style="font-size: small;">Changed write-host to write-output &amp; added "-NoProfile" to $TaskArg (thanks Pat)</span> </li>
</ul>
<p><span style="font-size: small;">v1.0 : 16th January 2015:</span></p>
<ul>
<li><span style="white-space: pre;">&nbsp;</span><span style="font-size: small;">This is the original release</span> </li>
</ul>
<p>&nbsp;</p>
<h3><span style="font-size: small;">Credits</span></h3>
<p><span style="font-size: small;">Drago Totev for the original solution: <a title="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html" href="http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html"> http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html</a></span></p>
<p><span style="font-size: small;">Creating a task in PowerShell: <a href="http://www.verboon.info/2013/12/powershell-creating-scheduled-tasks-with-powershell-version-3/"> http://www.verboon.info/2013/12/powershell-creating-scheduled-tasks-with-powershell-version-3/</a></span></p>
<p><span style="font-size: small;">Tricky Task creation: <a href="http://stackoverflow.com/questions/20108886/scheduled-task-with-daily-trigger-and-repetition-interval"> http://stackoverflow.com/questions/20108886/scheduled-task-with-daily-trigger-and-repetition-interval</a> </span><br /> <span style="font-size: small;"><a href="https://p0w3rsh3ll.wordpress.com/2013/07/05/deprecated-features-of-the-task-scheduler/">https://p0w3rsh3ll.wordpress.com/2013/07/05/deprecated-features-of-the-task-scheduler/</a></span></p>
<p>&nbsp;</p>
<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/new-lyncmeetingwarmup/](https://greiginsydney.com/new-lyncmeetingwarmup/).
