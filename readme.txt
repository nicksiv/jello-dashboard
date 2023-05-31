Whatever happened to Jello Dashboard
=====================================

For quite some time now, I feel the need to explain my decision to stop developing the only widespread
application have ever developed.
Jello Dashboard was a great ride It was the best way for me to learn to do stuff l have never done before, like
maintaining a code base, reply to user comments, fixing bugs, building product web pages and a hundred more
smaller things, a person must do in a one-man application.
l started on my own to fulfill my own needs, but to my surprise, one after the other, many people showed
interest in what was building and helped me out in many different ways, something l will forever be grateful for
Arter almost 10 years of development, the project had to come to an end for several reasons. The main building
blocks of the application, JavaScript and HTML, were just using a feature of Outlook (the folder homepage) which
was exposing the outlook object model to many risks, so Microsoft applied numerous security fixes limiting my
application s functionality l worked around methods to bypass some of the restrictions, but in the end l was
feeling like a virus developer who tries to exploit security weaknesses
Apart from that, we are now in the mobile sync era, users demand such features and applications which cannot
implement them, are bound to die slowly
Although I tried to synchronize my application withsome web services, the outcome was not successful and
without modern debugging capabilities (l used a simple text editor and IE debugger) l couldn't cope with such
complexity.
And that's about it. Thanks to all the people who tested and used my application, who submitted code, did
translations and gave life to my project.


Jello Dashboard 5.30 beta
-----------------------------------------------------------------
This application is released under the Open Source GPLv3 license.
=================================================================

Jello Dashboard 5 is a complete GTD management system for Outlook.
Although it is considered to be an addin its more than a custom homepage for Outlook folders.

Features
--------
-  Easily organize your actions in lists using tags
-  File tags into folders using a system hierarchy 
-  Use special tags for project management
-  Manage Inbox quickly and keep it empty 
-  Manage reference items and connect them to your actions
-  Collect information in bulk from folders or by free text writing 
-  Work with Ticklers from the Outlook Calendar and Task due dates 
-  See all information important to you using the homepage widgets 
-  Review your action lists easily 
-  Get your lists, print or send them 
-  Work with your familiar Outlook application and its items 
-  Control application’s behavior through an extensive set of user settings 

Jello Dashboard will load when you select you folder of choice from the Outlook folder hierarchy. 

After extracting the contents of this file to your folder of choice, please follow the manual Outlook folder assignment procedure describing below:

1. Navigate to your Outlook folder structure and select the folder in which you want Jello to be installed.
2. Right click on the folder and select properties.
3. Click on the Homepage tab button at the top.
4. Click the browse button and navigate to the folder you have previously extracted the Jello 5 files.
5. Select the jello5.htm and press OK select the file.
6. Check the Show home page by default for this folder checkbox and press OK.

Now everytime you click on your Outlook folder of choice in which you have installed Jello 5 the application will load.


To uninstall, just go to the Outlook folder properties again and uncheck the Show home page by default for this folder to view your previously used Outlook view for the folder.

*Note that Outlook 2007 users and above who want to install Jello Dashboard as a folder homepage can only use their default store. If the default store is on an Exchange server only an Exchange server Public folder can be used.

Alternatively you can try to run jello.hta standalone or jello5.htm file directly in Internet explorer or Firefox using the IETab plugin.

=================================================================


Nicolas Sivridis
https://jellodashboard.blogspot.com/
