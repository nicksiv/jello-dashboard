//    This file is part of Jello Dashboard.
///
//    Jello Dashboard is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.

//    Jello Dashboard is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with Jello Dashboard.  If not, see <http://www.gnu.org/licenses/>.
//
//    2008-2012 N.Sivridis http://jello-dashboard.net

//-----settings engine defaults------
var useTables=true;

//----define global variables----
var jelloVersion="Jello.Dashboard 5.26 beta (Astral)";
var itms=new Array();
var jello;
var its;
var catProperty=8;							//Category property id
var dueProperty=36;							// DueDate property ID
var alldayProperty=29;						//AllDayEvent property ID
var lastContext="";
var conStatus=null;
var thisProgress=0;
var lastOpenTagID=0;
var jese;
var panelWidth=0;
var panelHeight=0;
var thisGrid=null;
var migFols=null;
var history=new Array();
var historyPlace=0;
var addFolderTemp="";
var addFolderNd="";
var defaultWidgetHeight=170;
var elmRemoveTag="X";
var webJDversion="";
var webCheckHours=240; //check for updates every 10 days
var setExportFile="";

//Icons
var imgPath="img\\";
var icContact="contact.gif";
var icTask="task.gif";
var icTaskP="taskprv.gif";
var icTaskNS="taskns.gif";
var icApp="appoint.gif";
var icMail="mail.gif";
var icPost="post.gif";
var icProject="project.gif";
var icIsNext="isnext.gif";
var icNoNext="nonext.gif";
var icCheck="check.gif";
var icNote="note.gif";
var icNotes="notes.gif";
var icLinks="link.gif";
var icJournal="journal.gif";
var icFolder="folder.gif";
var icTag="list_components.gif";
var icMng="mng.gif";
var icMnger="mngC.gif";
var icFlag="flag.gif";
var gotoHasTags=false;
var dpass = new Array();
var jelloNotesDelim = "<-- " + txtJelloNotes + " -->";
var jelloStepsDelim = "<-- " + txtJelloSteps + " -->";
var OL2007VwMargin=30;
var isCollectScreen=false;
var firstLoadHome=true;

function jelloSettingsDefaults()
{
this.appTitle="Jello.Dashboard 5 (Valis)";
this.appLanguage="en";
this.inboxFolder=inboxItems.EntryID;
this.actionFolder=taskItems.EntryID;
this.contactFolder=contactItems.EntryID;
this.calendarFolder=calendarItems.EntryID;
this.archiveFolder="";
this.updateOutlookCategories=true;
if (OLversion>=12){this.autoUpdateSecuredFields=true;}else{this.autoUpdateSecuredFields=false;}
this.mailPreviewHTML=true;
this.tagCounters=false;
this.OLCounters=false;
this.sidebarLeft=false;
this.newActionStatus=1;
this.showNotStartedAlways=true;
this.dateFormat=0;
this.timeFormat=0;
this.dateSeperator="/";
this.dueDisplayDaysLeft=false;
this.dueShowYear=false;
this.agendaDays=10;
this.hidePrivateApps=false;
this.calendarType=0;
this.defaultCalendarView="0";
this.CalendarShowCompleted=true;
this.pageSize=10;
this.onlyTaskInbox=false;
this.alertSeconds=5;
this.migrated=false;
this.removeLinked=0;
this.autoUpdateTaskNotes=false;
this.tags=[{"id":1,"parent":0,"tag":txtTags,"istag":false,"notes":"","archive":"","color":"","icon":"","isdoc":false,"private":false,"isproject":false},{"id":2,"parent":1,"tag":"@Home","istag":true,"notes":"Things to do at home","private":false,"isproject":false},{"id":3,"parent":1,"tag":"@Work","istag":true,"notes":"Things to do at work","private":false,"isproject":false},{"id":5,"parent":0,"tag":txtSystem,"istag":false,"notes":"System tags","private":false,"sys":true,"isproject":false},{"id":6,"parent":5,"tag":"!Next","istag":true,"notes":txtNextActions,"private":false,"sys":true,"isproject":false},{"id":7,"parent":5,"tag":"@Review","istag":true,"notes":"Items for review","private":false,"sys":true,"isproject":false},{"id":8,"parent":5,"tag":"@Waiting","istag":true,"notes":"Waiting for","private":false,"sys":true,"isproject":false},{"id":9,"parent":5,"tag":"@Someday","istag":true,"notes":"Someday Maybe","private":false,"sys":true,"isproject":false},{"id":10,"parent":0,"tag":txtProjects,"istag":false,"notes":"Store your projects here","private":false,"isproject":false}];
this.lastTagId=10;
this.shortcuts=[{"id":1,"shortcut":txtNextActions,"cmd":"showContext(6,1);","sys":true,"icon":"isnext.gif"},{"id":2,"shortcut":txtWaiting,"cmd":"showContext(8,1);","sys":true,"icon":"list_waiting.gif"},{"id":3,"shortcut":txtReview2,"cmd":"showContext(7,1);","sys":true,"icon":"list_review.gif"},{"id":4,"shortcut":txtSomeday,"cmd":"showContext(9,1);","sys":true,"icon":"list_someday.gif"}];
this.outlookFolders=[];
this.lastShortcutId=4;
this.importSettings=[];
this.lastImportSetId=0;
this.widgetColumns=2;
this.widgets=[];
this.wGridTextSize=9;
this.gridTextSize=11;
this.gridTextFamily="Verdana";
this.gridColumns=[];
this.mailPreviewHeight=180;
this.actionPreviewHeight=180;
this.mailPreviewFontSize=10;
this.hidePastAppointments=false;
this.collectAdvancedMode=false;
this.hideOLPreviewPane=false;
this.inboxAutoRefresh=false;
this.ticklerPopupDuration = 14;
this.reviewShowEmptyTags=false;
this.weeklyListDefs="80|10|"+txtNotes+"|@Waiting|";
this.theme=0;
this.markReadfromJello=true;
this.useOLDue = false;
this.sendTaskRequests = false;
this.useGoToFolderToMove=true;
this.addSpaceOlView=false;
this.startup=0;
this.activeSBPanel=1;
this.sidebarOn=true;
this.sidebarWidth=200;
this.firstDayWeek=1;
this.firstRun=1;
this.selectFirstItem=0;
this.nextToggleHighPri=0;
this.reviewPanelPaths=0;
this.nonActionFolder=journalItems.EntryID;
this.noUseMail=0;
this.goNextOnDelete = true;
this.doubleClickBodyPopup = false;
this.enableBccProcess=false;
this.bccCategory = "";
this.bccMatch="";
this.reviewpopup=0;
this.allODueTasks=1; 
this.previewState=1;
this.ALinlineduedate=0;
this.expBodyView=0;
this.enablekeys03=0;
this.lastWebVersionCheck="";
this.createAppointmentsOnDefCal=1;
this.useTables=false;
this.autoWebVersionCheck=true;
this.htaWidth=600;
this.htaHeight=600;
this.inboxGroupOn=true;

}



function jelloSettings()
{
//jello settings object
def=new jelloSettingsDefaults;
this.journalEntryCreated=false;
  
  this.saveCurrent=function()
  {
  //save current settings
  var jsonStr=Ext.util.JSON.encode(jello);
	this.save(jsonStr);
  };

	this.restoreDefaults=function()
	{
	//add default values to settings
	var jsonStr=Ext.util.JSON.encode(def);
	this.save(jsonStr);
	};

	this.save=function(ss)
	{
	//save setting string
  try {var test=this.systemJournalEntry.UserProperties("settingString").Value;}
	catch(e){this.systemJournalEntry.UserProperties.Add("settingString",1,0);}
		this.systemJournalEntry.UserProperties("settingString").Value=ss;
		this.systemJournalEntry.Save();

	};

 this.getSettingString=function()
 {
 //get setting string from the settings journal entry

 //var ss="";
 try{var ss=this.systemJournalEntry.UserProperties("settingString").Value;
 return ss;
}catch(e){return null;}

 };

 var lookupJSet=false;
 if (!notEmpty(jelloSettingsFolder)){jelloSettingsFolder=journalItems.EntryID;}

    try{var jfldr=NSpace.GetFolderFromID(jelloSettingsFolder);}
	catch(e)
	{
		//cannot load custom journal entry folder. Enumerate to force loading
	   document.title ="Forcing loading stores...";
	   var cStores = NSpace.Folders;
		for( var i=1; i<=cStores.Count; i++)
		{
		 try{ var aCol = cStores(i); var cc=aCol.Folders.GetFirst();document.title ="Loading store "+aCol+"...";}
		 catch(f){}
		}
	//alert("Apparently your custom folders could not be loaded correctly because you are running Jello Dashboard outside Outlook.\nPlease expand the Outlook store which contains your settings journal entry and Cancel the following dialog.\nPress OK to see the folders list.");
	//var xxx= NSpace.PickFolder();
	}

	 do
	 {
	     try
			  {
				var jfldr=NSpace.GetFolderFromID(jelloSettingsFolder);
			}
		catch(e)
		{

		 lookupJSet=confirm("Settings journal folder could not be loaded. \nDo you want to select it yourself?");
			if (lookupJSet)
			{var nf=NSpace.PickFolder();
			if (!notEmpty(nf)){lookupJSet=true;return;}
			else{jfldr=nf;lookupJSet=false;}
			}
		}
	 }while(lookupJSet==true);


  try
  {
  var jfldr=NSpace.GetFolderFromID(jelloSettingsFolder);
  }
  catch(e)
  {
  alert("Set folder not found. Will search to the default journal folder.");
  var jfldr=journalItems;jelloSettingsFolder=journalItems.EntryID;
  }

  var je=jfldr.Items;
	var j=je.Restrict("[Subject]='"+jelloSettingsName+"'");

    if (j.Count==0)
		{
		var ja=je.Add();
		ja.subject=jelloSettingsName;
		ja.body=txtSettingBody;
		this.journalEntryCreated=true;
		if (ja.UserProperties.Count==0){ja.UserProperties.Add("settingString",1,0);}
		ja.Save();
		this.systemJournalEntry=ja;
		}
	    else
	    {
      this.systemJournalEntry=j(1);

        if (this.systemJournalEntry.UserProperties.Count==0){this.journalEntryCreated=true;}
        else
        {if(!notEmpty(this.systemJournalEntry.UserProperties("settingString").Value)){this.journalEntryCreated=true;}}
      }
	    if (this.journalEntryCreated==true){this.restoreDefaults();}
	    var gss=this.getSettingString();
	    if (gss==false){var gss=this.getSettingString();}

		  jello=Ext.util.JSON.decode(gss);

}



function progressBar(m)
{	//**j5
var br=thisProgress;
thisProgress--;
var e="starting ";
if (m!=null){e=m;}
for(var y=0;y<br;y++){e+=" | ";}
status=e;
}

function startIntTimer()
{  //**j5
	// get a tick every 15 seconds
	homeTimerID = setInterval("updateMainDate();", 15000);
}

function tagExists(tg)
{
tagStore.filter("tag",tg);
var cc=tagStore.getCount();
tagStore.clearFilter();
if (cc==0){return false;}else{return true;}
}


function recordToJSON(vr,tr,vs)
{//add an ext record to json obj variable
var str=Ext.util.JSON.encode(vr);
var str2=Ext.util.JSON.encode(tr.data);
var y=vs+"="+str.substr(0,str.length-1)+","+str2+"];";
eval(y);
}

function updateJSONitems(ds){
  var buff = [];
  var toJSON = function(record){
      buff.push (Ext.util.JSON.encode(record.data ));
  };
   ds.each(toJSON);
   return '[' + buff.join(',') + ']';
 };

function syncStore(theStore,vr)
{//synchronize store to variable
var y=updateJSONitems(theStore);
eval(vr+"="+y);
}

function init()
{
status="Connecting...";
var con=connect();
if (con==false){return false;}

jese=new jelloSettings();
    //check for undefined settings
    for (prop in def)
    {
    if(typeof(jello[prop])=="undefined")
    {
    jello[prop]=def[prop];
    jese.saveCurrent();
    }
	}


if (jello.appLanguage!="en"){setLang(jello.appLanguage);}
    setTimeout(function(){
    if (jese.journalEntryCreated){Ext.info.msg("Jello Settings","Jello Settings journal entry created");}
    },100);



if (jello.useTables==0 || jello.useTables=="0"){useTables=false;}
if( OLversion >= 12 && useTables){inboxTable = getTable(inboxItems,"");}

return true;


}

function setLang(lg)
{
//function to be used for language setting
langscript.src="langs\\lang-"+lg+".js";
}

function setTheme(tm)
{
//set the selected extjs theme
//
if (tm=="1" || tm==1 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-gray.css');}
if (tm=="2" || tm==2 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-silverCherry.css');}
if (tm=="3" || tm==3 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-midnight.css');}
if (tm=="4" || tm==4 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-blueen.css');}
if (tm=="5" || tm==5 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-yellow.css');}
if (tm=="6" || tm==6 ){Ext.util.CSS.swapStyleSheet('themecss', 'extjs\\resources\\css\\xtheme-pink.css');}
}

function validateAllFolders()
{
//check if all folders reffered in settings are valid
//Does not check inbox folder for now. It can be an IMAP offline. Check others.
//var a1=setAndCheckArcFolder(jello.inboxFolder);
var a2=setAndCheckArcFolder(jello.actionFolder);

	   if (a2==null){
     document.title ="Forcing loading stores...";
	   var cStores = NSpace.Folders;
		for( var i=1; i<=cStores.Count; i++)
		{
		 try{ var aCol = cStores(i); var cc=aCol.Folders.GetFirst();document.title ="Loading store "+aCol+"...";}
		 catch(f){}
		}
		 }

var a2=setAndCheckArcFolder(jello.actionFolder);
var a3=setAndCheckArcFolder(jello.contactFolder);
var a4=setAndCheckArcFolder(jello.calendarFolder);

var msg="";
if (a2==null){jello.actionFolder=taskItems.EntryID();msg+="Action ";}
if (a3==null){jello.contactFolder=contactItems.EntryID();msg+="Contacts ";}
if (a4==null){jello.calendarFolder=calendarItems.EntryID();msg+="Calendar ";}

  if (notEmpty(msg))
  {
    var cmsg=txtFolderValidMsg.replace("%1",msg);
    var choice=confirm(cmsg);
    if (choice){jese.saveCurrent();location.reload;}
  }
}

function updateCSS()
{
//update stylesheet to user's prefs
Ext.util.CSS.updateRule(".x-grid3-cell-inner","font-size",jello.wGridTextSize);
Ext.util.CSS.updateRule(".x-grid3-cell-inner","font-family",jello.gridTextFamily);
Ext.util.CSS.updateRule("#grid .x-grid3-cell-inner","font-size",jello.gridTextSize);
Ext.util.CSS.updateRule("#mgrid .x-grid3-cell-inner","font-size",jello.gridTextSize);
Ext.util.CSS.updateRule("#igrid .x-grid3-cell-inner","font-size",jello.gridTextSize);
Ext.util.CSS.updateRule("#tgrid .x-grid3-cell-inner","font-size",jello.gridTextSize);
Ext.util.CSS.updateRule("#grid .x-grid3-cell-inner","font-family",jello.gridTextFamily);
Ext.util.CSS.updateRule("#mgrid .x-grid3-cell-inner","font-family",jello.gridTextFamily);
Ext.util.CSS.updateRule("#tgrid .x-grid3-cell-inner","font-family",jello.gridTextFamily);
Ext.util.CSS.updateRule("#igrid .x-grid3-cell-inner","font-family",jello.gridTextFamily);
}

//start the application
var initOK=init();

Ext.EventManager.on(window, 'beforeunload', function(){

//restoreOutlookDefaults();
});

Ext.EventManager.on(window, 'unload', function(){

});



//Auto startup to ...
Ext.EventManager.on(window, 'load', function(){
if (initOK){
try{updateWindowSize();}catch(e){}
try{
checkVersion();}catch(e){}

validateAllFolders();
updateCSS();


if(jello.actionPreviewHeight>Ext.getBody().getViewSize().height){jello.actionPreviewHeight=200;jese.saveCurrent();};

if (jello.theme!="0" && jello.theme!=0 ){setTheme(jello.theme);}
if (jello.firstRun==1){createDonationAction();}


  setTimeout(function(){

  try{
  if (jello.startup==0){pHome();}
  if (jello.startup==3){pHome();}
  if (jello.startup==1){pInbox();}
  }catch(e){}

try{setSidebarPanelHeights();}catch(e){}

  },3);

  setTimeout(function(){
try{Ext.getCmp("previewpanel").update(smallAbout());}catch(e){}
status=txtJDLoaded;
},6);


  }
});

Ext.EventManager.on(window, 'resize', function(){
htaResize();

});

function checkVersion(force)
{

var td=new Date();

if(notEmpty(jello.lastWebVersionCheck))
{
  var ld=new Date(jello.lastWebVersionCheck);
  var uhours=hourDif(td,ld);
  if (uhours<webCheckHours && !force){return;}
}
var rurl="http://jello-dashboard.net/downloads/jdversion.txt";

	if (force==null)
	{
		if (jello.autoWebVersionCheck==false || jello.autoWebVersionCheck=="0" || jello.autoWebVersionCheck==0)
		{return;}
		
	}

try{
Ext.Ajax.request({
   url: rurl,
   success:function(a,b,c)
   {
   webJDversion=a.responseText;
   if (webJDversion!=jelloVersion)
   {
   
   try{Ext.info.msg("Application update","A new Jello Dashboard version is available.<br>"+webJDversion+"<br><a href='http://www.jello-dashboard.net/download/' target='_blank'><b>Go to the downloads page</b></a>");}catch(e){}
   try{Ext.getCmp("previewpanel").update(smallAbout());}catch(e){}
   jello.lastWebVersionCheck=td.toUTCString();
   jese.saveCurrent();

   }
   else
   {
   if (force){Ext.info.msg("Application update","You have the latest version");}
   }
   }

});}catch(e){}


}