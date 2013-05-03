//    This file is part of Jello Dashboard.
//
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

var sitemCount=0;
var settingReloadFlag=false;


function mngSettings(onPage){
settingReloadFlag=false;

initScreen(false,"mngSettings("+onPage+")");
	
thisGrid="portalpanel";
//main.innerHTML="<div id=toolbar></div><div id=msg></div>";

//render toolbar
var menu2 = new Ext.menu.Menu({
        id: 'utilmenu',
        items: [{
        icon: 'img\\icon_extension.gif',
        cls:'x-btn-icon',
        hidden:olcatshidden,
        text: txtUpdateOutlookCategoriesNow,
			listeners:{
            click: function(b,e){updateOLCatAll();}
            }},{
        icon: 'img\\page_package.gif',
        cls:'x-btn-icon',
        text: txtMigHomeMsg,
			listeners:{
                click: function(b,e){
                dynamicLoad("migrate.js","settingsMigration()");
                }
            }
        },{
        icon: 'img\\setimp.gif',
        cls:'x-btn-icon',
        text: txtSIImpAF,
			listeners:{
            click: function(b,e){importAllOLCategories();}
            }
        },{
        icon: 'img\\setimp.gif',
        cls:'x-btn-icon',
        text: txtSIImpAFAny,
			listeners:{
            click: function(b,e){importAllOLCategories(true);}
            }
        },
		{icon: 'img\\setimp.gif',
        cls:'x-btn-icon',
        text: txtImportAllOLC,
			listeners:{
            click: function(b,e){importNSpaceOLcategories();}
            },
		hidden:(OLversion >= 12?false:true)
		},
		{
        icon: 'img\\sfolder.gif',
        cls:'x-btn-icon',
        text: txtSIFixOrphans,
			listeners:{
            click: function(b,e){fixOrphanTags();}
            }
        }

        ]}
    );

    var olcatshidden=true;
if (OLversion>10){olcatshidden=false;}
 stbar = new Ext.Toolbar();
//    stbar.render('toolbar');
        stbar.add({
        icon: 'img\\action_save.gif',
        cls:'x-btn-icon',
        tooltip: txtSIBackup,
			listeners:{
            click: function(b,e){backupSettings();}
            }
        },'-',{
        icon: 'img\\icon_alert.gif',
        cls:'x-btn-icon',
        tooltip: txtSIReset,
			listeners:{
            click: function(b,e){resetSettings();}
            }
        },{
        icon: 'img\\mail.gif',
        cls:'x-btn-icon',
        tooltip: txtSISend,
			listeners:{
            click: function(b,e){sendSettings();}
            }
        },{
        icon: 'img\\page_package.gif',
        cls:'x-btn-icon',
        tooltip: txtSIView,
			listeners:{
            click: function(b,e){viewSettings();}
            }
        },'-',{
            text:txtMenuOtherFunctions,
            menu: menu2
        });

 var tabs = new Ext.TabPanel({
  //      renderTo: main,
        id:"tbs",
        tbar:stbar,
        activeTab: onPage,
        autoHeight:true,
        autoScroll: true,
        listeners:{tabchange:function(){resizeThings();}},
        plain:true,
        defaults:{autoScroll: true},
            items:[
                new Ext.FormPanel({
                title:txtSetDefaults,
                //height:500,
                id:'sdef',
                buttonAlign:'right',
                labelWidth:250,
                autoHeight:true,
				autoScroll: true,
                frame:true,
                items:[
                new Ext.form.FieldSet({id:'gfolders',title:txtSetFolders,collapsible:true,autoHeight:true}),
                new Ext.form.FieldSet({id:'gaction',title:txtCbActions,collapsible:true,autoHeight:true}),
                new Ext.form.FieldSet({id:'gmail',title:txtSetMessages,collapsible:true,autoHeight:true}),
                new Ext.form.Hidden({	})
                ]
        }),
                new Ext.FormPanel({
                title:txtSetCore,
                autoHeight:true,
                id:'score',
                labelWidth:250,
                buttonAlign:'right',
                listeners:{render: function(){
                setTimeout(function(){Ext.getCmp("ggrid").doLayout();Ext.getCmp("ggui").doLayout();},3);}},
                frame:true,
                items:[
                new Ext.form.FieldSet({id:'ggui',title:txtSetGUI,collapsible:true,autoHeight:true}),
                new Ext.form.FieldSet({id:'ggrid',title:txtSetGrid,collapsible:true,autoHeight:true}),
                new Ext.form.Hidden({	})
                ]

        }),
         new Ext.FormPanel({
                title:txtCalendar,
                autoHeight:true,
                id:'scal',
                buttonAlign:'right',
                labelWidth:250,
                frame:true,
                items:[
                new Ext.form.Hidden({	})
                ]
        }),
        new Ext.FormPanel({
                title:'Experimental',
                autoHeight:true,
                id:'sexp',
                buttonAlign:'right',
                labelWidth:350,
                frame:true,
                items:[
                new Ext.form.Hidden({	})
                ]
        }),

              
        
        {
				    title: txtSetInfo,
				    frame:true,
				    html: genSettings(),
				    autoHeight:true
			     }
        ]
    });




    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(tabs);
    ppnl.setAutoScroll(true);
ppnl.setTitle("<img src=img\\settings16.png style=float:left;>"+txtSettings);

addSettingItems();
//**************
setTimeout(function(){resizeGrids();ppnl.doLayout();},136);
}

function addSettingItems(){
var q=Ext.getCmp("gfolders");
var q1=Ext.getCmp("gaction");
var q2=Ext.getCmp("gmail");
var q3=Ext.getCmp("ggrid");
var q4=Ext.getCmp("ggui");

printListSetting(3,jello.inboxFolder,"inboxFolder","",txtSetInboxFolder,icFolder,q);
printListSetting(0,jello.noUseMail,"noUseMail","",txtSetNoUseMail,icMail,q2);
printListSetting(3,jello.actionFolder,"actionFolder","",txtSetActionItems,icTask,q);

//implementation attempt for the todo bar items in newer OL versions
//printListSetting(5,"","actionFolder2","","Use Todo Items",icTask,q);

printListSetting(3,jello.archiveFolder,"archiveFolder","",txtSetArchiveFolder,icFolder,q);
printListSetting(3,jello.contactFolder,"contactFolder","",txtSetContactFolder,"contact.gif",q);
printListSetting(3,jello.nonActionFolder,"nonActionFolder","",txtSetNonActionFolder,icJournal,q);

printListSetting(0,jello.autoUpdateSecuredFields,"autoUpdateSecuredFields","",txtAutoUpdateSecuredFields,icFolder,q2);

printListSetting(0,jello.mailPreviewHTML,"mailPreviewHTML","",txtMailPreviewHTML,icFolder,q2);
printListSetting(0,jello.autoUpdateTaskNotes,"autoUpdateTaskNotes","",txtAutoUpdateNotes,icTask,q1);
printListSetting(2,jello.removeLinked,"removeLinked","0;"+txtSetLiItmsOnDel1+";1;"+txtSetLiItmsOnDel2+";2;"+txtSetLiItmsOnDel3+";",txtremoveLinked,icFolder,q1);
printListSetting(2,jello.newActionStatus,"newActionStatus","1;"+txtInProgress+";0;"+txtNotStarted,txtSNewTaskDefault,icTask,q1);
printListSetting(0,jello.showNotStartedAlways,"showNotStartedAlways","",txtSetShowNotStarted,icTask,q1);
printListSetting(0,jello.nextToggleHighPri,"nextToggleHighPri","",txtnextToggleHighPri,icIsNext,q1);
printListSetting(0,jello.sendTaskRequests,"sendTaskRequests","",txtSetTaskRequests,"user.gif",q1);

printListSetting(0,jello.selectFirstItem,"selectFirstItem","",txtSelectFirstLst,icNotes,q3);
printListSetting(2,jello.gridTextFamily,"gridTextFamily",txtGridFontChoices,txtgridTextFamily,icNotes,q3);
printListSetting(4,jello.gridTextSize,"gridTextSize","",txtGridTextSize,icNotes,q3);
printListSetting(4,jello.wGridTextSize,"wGridTextSize","",txtWGridTextSize,icNotes,q3);
printListSetting(0,jello.enablekeys03,"enablekeys03","",txtEnablekeys03,icNotes,q3);
printListSetting(4,jello.pageSize,"pageSize","",txtSetPageSize,icFolder,q3);
printListSetting(0,jello.addSpaceOlView,"addSpaceOlView","",txtaddSpaceOlView,icFolder,q3);

if (OLversion>10)
{
printListSetting(0,jello.updateOutlookCategories,"updateOutlookCategories","",txtUpdateOutlookCategories,icNotes,q2);
}

printListSetting(1,jello.appTitle,"appTitle","",txtSetTitle,icNotes,q4);
printListSetting(2,jello.appLanguage,"appLanguage",langList(),txtSetLang,icNotes,q4);
printListSetting(2,jello.theme,"theme","0;Default;1;Gray;2;Silver Cherry;3;Midnight;4;Blueen(Javier Rincon);5;Yellow;6;Pink;",txtSetTheme,icNotes,q4);
printListSetting(2,jello.startup,"startup","0;"+txtHome+";3;"+txtHomeMiniz+";1;"+txtInbox,txtStartup,icNotes,q4);
printListSetting(4,jello.alertSeconds,"alertSeconds","",txtAlertSeconds,icNotes,q4);
printListSetting(0,jello.sidebarLeft,"sidebarLeft","",txtSetSidebarLeft,icNotes,q4);
//printListSetting(0,jello.onlyTaskInbox,"onlyTaskInbox","",txtonlyTaskInbox,icTask,q4);
printListSetting(0,jello.inboxAutoRefresh,"inboxAutoRefresh","",txtinboxAutoRefresh,"refresh.gif",q2);
printListSetting(0,jello.useGoToFolderToMove,"useGoToFolderToMove","",txtUseGoToFolderToMove,"move.gif",q2);
if(OLversion >= 12){
printListSetting(0,jello.useTables,"useTables","","Use Outlook Tables for faster performance<br>(valid antivirus required)",icNotes,q4);
}
printListSetting(0,jello.autoWebVersionCheck,"autoWebVersionCheck","","Auto-check for version changes every 10 days",icNotes,q4);

printListSetting(0,jello.markReadfromJello,"markReadfromJello","",txtmarkReadfromJello,"mail.gif",q2);
printListSetting(0,jello.goNextOnDelete,"goNextOnDelete","",txtgoNextOnDelete,"mail.gif",q2);

//if (OLversion>10){printListSetting(0,jello.hideOLPreviewPane,"hideOLPreviewPane","",txthideOLPreviewPane,"olpreview.gif",q2);}

var p=Ext.getCmp("scal");
printListSetting(3,jello.calendarFolder,"calendarFolder","",txtSetCalendarFolder,icApp,p);
printListSetting(0,jello.createAppointmentsOnDefCal,"createAppointmentsOnDefCal","","Create Appointments on Default Calendar Only","notes.gif",p);

var jds=jello.dateSeperator;
printListSetting(2,jello.dateFormat,"dateFormat","0;d"+jds+"m"+jds+"y;1;m"+jds+"d"+jds+"y;2;yyyy"+jds+"mm"+jds+"dd",txtDateFormat,icApp,p);
printListSetting(2,jello.timeFormat,"timeFormat","0;24 hrs;1;12 hrs",txtTimeFormat,icApp,p);
printListSetting(1,jello.dateSeperator,"dateSeperator","",txtDateSeparator,icApp,p);
var dll=txtDayList.split(",");

printListSetting(2,jello.firstDayWeek,"firstDayWeek","0;"+dll[0]+";1;"+dll[1],txtFirstDayWeek,icApp,p);
printListSetting(0,jello.dueDisplayDaysLeft,"dueDisplayDaysLeft","",txtSetDaysLeft,icApp,p);
printListSetting(0,jello.dueShowYear,"dueShowYear","",txtSDDShowYear,icApp,p);
printListSetting(0,jello.hidePastAppointments,"hidePastAppointments","",txtHidePastDue,"calendar.gif",p);
printListSetting(0,jello.CalendarShowCompleted,"CalendarShowCompleted","",txtCalendarShowCompleted,icApp,p);
printListSetting(4,jello.agendaDays,"agendaDays","",txtSetAgenda,icApp,p);
printListSetting(4,jello.ticklerPopupDuration,"ticklerPopupDuration","",txtTicklerPopupDuration,"calendar.gif",p);

//printListSetting(0,jello.hidePrivateApps,"hidePrivateApps","",txtSHidePrivateApps,icApp,p);
printListSetting(2,jello.defaultCalendarView,"defaultCalendarView","0;"+txtDay+";1;"+txtWeek+";2;"+txtMonth+";3;"+txtAgenda,txtSDefCalView,icApp,p);
printListSetting(0,jello.allODueTasks,"allODueTasks","",txtAllODueTasks,icTask,p);

//experimental
var p=Ext.getCmp("sexp");
if(OLversion >= 12){
printListSetting(0,jello.useOLDue,"useOLDue","",txtUseOLDue,icApp,p);
}
printListSetting(0,jello.doubleClickBodyPopup,"doubleClickBodyPopup","",txtdoubleClickBodyPopup,"mail.gif",p);
printListSetting(0,jello.enableBccProcess,"enableBccProcess","",txtEnableBccProcess+ " <a class=jellolinktop onclick=bcchelp()>[help]</a>","icon_extension.gif",p);
printListSetting(1,jello.bccCategory,"bccCategory","",txtBccCategory,"notes.gif",p);
printListSetting(1,jello.bccMatch,"bccMatch","",txtBccMatch,"notes.gif",p);
printListSetting(0,jello.reviewpopup,"reviewpopup","",txtReviewpopup,"list_someday.gif",p);
printListSetting(0,jello.ALinlineduedate,"ALinlineduedate","",txtALinlineduedate,"calendar.gif",p);
printListSetting(0,jello.expBodyView,"expBodyView","",txtExpBodyView,"icon_monitor_pc.gif",p);


setTimeout(function(){q.doLayout();
q1.doLayout();
q2.doLayout();
updateUseOfMail(jello.noUseMail);
},3);

}

function bcchelp()
{
var ret=txtBcchelp;
var ret="A new, experimental feature has been added to Jello 5.2 - process sent. <br>The purpose of this capability is automatically create tasks from certain sent items.  When sending an email if you include yourself (or a list of certain addresses) as a bcc the 'process sent' capability will find the messages and turn them into tasks for a category you specify.  The following describes the capability and the setting you must specify to use.";
ret+="<br><br><b>Settings:</b>";
ret+="<br><br>Enable Bcc task processing:  Check this box to enable the new 'Process Sent' capability.  After checking or unchecking this box you have to restart Jello (click the refresh icon) to get the function link to appear. Category to assign for bcc processing created tasks:  This is the list of categories that will be assigned to the items that are created.<br>Specify category names, separated by a semi-colon (;)/nEnter emails and user names to match in bcc field, semicolon separated:  This field MUST be filled in correctly.  Here you enter the email addresses that the 'process sent' command should look for in the bcc field of the sent items.  You may enter multiple email addresses, seprataed by semi-colons.  You MUST also include your user name.  When Outlook lets Jello look at email, we don't get the full email lists, but often get the email address converted to a user name of the contact with that email address.  You must enter that name as one of the values in this field.  You can obtain the value for this by looking at the name displayed in the 'Sender' column of items you send in the Sent Items folder, or send yourself an email and see what name appears in the inbox Sender column.";
ret+="<br><br><b>Using Process Sent</b>";
ret+="<br><br>First, to be able to have the 'Process Sent' capability find items to create tasks from, you must bcc yourself or some other address that will be matched in the settings described above.  When you want to process the sent items looking for these emails you click the link for 'Process Sent' that is below the search box on the right side of the Jello window.  If you haven't enabled the processing funciton the link will not be visible.  A prompt will appear asking you for the number of days to process.  The command will go back the specified number of days looking for items.  So if today is Tuesday and you specify '2', the command will look for any message after 12:00AM Sunday.  Specifying '0' processes messages for the current day (since 12:00AM today).  Leave this blank and the command is cancelled.";
ret+="<br><br>The 'Process Sent' command will now look at the Sent Items, find the matching emails, create tasks for the emails and assign the speficied categories to the created tasks.  After creating the tasks, the command will open a Review window for the category name that is first in the list.  So, if your list of categories is '@waiting;!next;@someday' the tasks will be assigned the 3 categories and the review view will open for the '@waiting' category.  Since no due date is assigned to the created tasks, they are easy to identify.";
ret+="<br><br><b>Preventing Repeat Processing</b>";
ret+="<br><br>The 'Process Sent' command tries to prevent duplicates.  When it processes a sent email and converts it to a task, the 'Mileage' filed of the sent items is modified to have a value.  Since this field is usually not used and in later versions of Outlook not even visible, it was thought safe to use this field for the purpose.";
ret+="<br><br><b>Mobile Devices</b>";
ret+="<br><br>On mobile devices using activesync (android, iphone, iPad, Windows phone), when you set a Bcc for yourself the activesync/exchange versions may not set this value on the Exchange server.  Try a test message from your device to see if the bcc field is copied or not.";
pLatest();
Ext.getCmp("latestform").update(ret);
Ext.getCmp("latestform").setTitle(txtBcchelpTitle);
Ext.getCmp("thelatestform").setTitle("Help");
}

function changeJelloSetting(el,val)
{
//change a jello setting
var id=el.getId();
if (val==true){val="1";}if (val==false){val="0";}
try{var ss=val.search("'");
if (ss>0){alert(txtMsgSetQuote);el.setValue(val.replace(new RegExp("'","g"),""));el.focus();return;}
}catch(e){}
var stt="jello."+id+"='"+val+"'";
eval(stt);

	if (id=="noUseMail")
	{
	updateUseOfMail(val);
	}

jese.saveCurrent();
checkForReload(id);
}

function checkForReload(id){
//some settings need the dashboard to be reloaded
	if (id=="appLanguage" || id=="tagCounters" || id=="OLCounters" || id=="sidebarLeft" || id=="theme" || id=="firstDayWeek")
	{
		if (settingReloadFlag==false)
		{


		Ext.Msg.show({
           title:txtWidTImportant,
           msg: txtReloadforEffect,
           buttons: {yes:txtReload,no:txtCancel},
           fn: function(b,t){if (b=='yes'){location.reload();}},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });

		}
	settingReloadFlag=true;
	}
}

function settingComboSelection(fld,rec)
{
var a=rec.get("value");
changeJelloSetting(fld,a);
}

function printListSetting(type,vr,cookie,lst,txt,ico,pnl,isnum)
{
//render a settings line type(0:checkbox 1:text 2:combo) vr(variable) cookie(setting id) last(value list for combos) txt(setting caption)
var bim="background-image:url("+imgPath+ico+");background-repeat:no-repeat;padding-left:18px;width:300;";
if (type==1)
	{
	try{vr=vr.replace(new RegExp("'","g"),"&#39;");}catch(e){}
  //textbox
	{
	var tfield = new Ext.form.TextField({
	fieldLabel:txt,
	value:vr,
	width:200,
	invalidText:txtInvalidTxt,
	listeners:
	{change:function(fld,nv,ov){changeJelloSetting(fld,nv);}},
	labelStyle:bim,
	hideLabel:false,
	id:cookie
	});
	}
}

if (type==4)
	{
	try{vr=vr.replace(new RegExp("'","g"),"&#39;");}catch(e){}
  //number textbox
	{
	var tfield = new Ext.form.NumberField({
	fieldLabel:txt,
	value:vr,
	width:200,
	invalidText:txtInvalidTxt,
	listeners:
	{change:function(fld,nv,ov){changeJelloSetting(fld,nv);}},
	labelStyle:bim,
	hideLabel:false,
	id:cookie
	});
	}
}

	if (type==0)
	{
	//checkbox
	var ck=false;if (vr==1){ck=true;}
	var tfield = new Ext.form.Checkbox({
	fieldLabel:txt,
	checked:ck,
	labelStyle:bim,
	listeners:
	{check:function(fld,nv){changeJelloSetting(fld,nv);}},
	hideLabel:false,
	id:cookie
	});

  }

  if (type==5)
	{
	//ToDo bar button
	var tfield = new Ext.Button({
	fieldLabel:txt,
	labelStyle:bim,
  text:'...',
	listeners:
	{click:function(fld,nv)
    {
    var tdbar=NSpace.GetDefaultFolder(28);
    var afff=Ext.getCmp("actionFolder")
    changeJelloSetting(afff,tdbar.EntryID);
    afff.setValue(tdbar.FolderPath);
    
    }
  },
	hideLabel:false,
	id:cookie
	});

  }
  
		if (type==2)
	{
	//combo
	 var comboRecord = Ext.data.Record.create([
    {name: 'value'},
    {name: 'text'}]);
  var comboStore=new Ext.data.SimpleStore({
        fields: [{name: 'value'},{name: 'text'}]
    });
  var l=lst.split(";");
  for (var x=0;x<l.length;x=x+2)
  {
    var cr=new comboValue(l[x],l[x+1]);
    var newRec=new comboRecord(cr);
    comboStore.add(newRec);
  }
  var tfield = new Ext.form.ComboBox({
	fieldLabel:txt,
	value:vr,
	triggerAction:'all',
	mode:'local',
	labelStyle:bim,
	width:200,
	listeners:
	{select:function(fld,rec,idx){settingComboSelection(fld,rec);}},
	displayField:'text',
	valueField:'value',
	editable:false,
	hideLabel:false,
	store:comboStore,
	id:cookie
	});
  }

	if (type==3)
	{//folder chooser
	var jf="";try{jf=NSpace.GetFolderFromID(vr);var jfp=jf.FolderPath;}catch(e){var jfp=txtUndefined;}
	var tfield = new Ext.form.TriggerField({
	fieldLabel:txt,
	value:jfp,
	width:300,
	labelStyle:bim,
	onTriggerClick:function(e){fchTriggerClick(this);},
	hideLabel:false,
	id:cookie
	});
  var vc="val"+cookie;
  var hfield = new Ext.form.Hidden({
	value:vr,
	id:vc,
	renderTo:pnl.body
	});
	}

	pnl.insert(sitemCount,tfield);
	///pnl.insert((sitemCount+1),hfield);
  sitemCount=sitemCount+1;
	}

function genSettings(){
//general settings
var ret="<br>";
var uri=getAppPath();
var jF=jese.systemJournalEntry.Parent;
var jMD=new Date(jese.systemJournalEntry.LastModificationTime);
ret+="<table align=center cellpadding=3 cellspacing=3 width=95%>";
ret+= "<tr><td class=printChars>"+txtVersion+":</td><td class=printChars>" + jelloVersion + "</td></tr>";
ret+= "<tr><td class=printChars>Outlook Version:</td><td class=printChars>" + OLversion + "</td></tr>";
ret+= "<tr><td class=printChars>ExtJS Version:</td><td class=printChars>" + Ext.version + "</td></tr>";
ret+= "<tr><td colspan=5>&nbsp;</td></tr>";
ret+= "<tr><td class=printChars>"+txtDashPath+":</td><td class=printChars colspan=5><a class=jellolinktop href='"+uri+"'><b>" + uri + "</b></a></td></tr>";
ret+= "<tr><td class=printChars>"+txtSetIFCon+":</td><td colspan=5 class=printChars>" + conStatus + "</a></td></tr>";
ret+= "<tr><td class=printChars>"+txtSetIFSUse+":</td><td class=printChars colspan=5><b>" + Ext.util.Format.htmlEncode(jelloSettingsName); + "</b> ("+txtSetIFLCh+" "+jMD.format('j F Y @H:i')+")</a></td></tr>";
ret+= "<tr><td class=printChars>"+txtSetIFSFol+":</td><td class=printChars colspan=5><a class=jellolinktop onclick=openSettingsFolder()><b>" + jF.FolderPath +"</b></a></td></tr>";
ret+= "<tr><td colspan=5>&nbsp;</td></tr>";
ret+= "<tr><td class=printChars>"+ScriptEngine()+" Version:</td><td class=printChars colspan=5><b>" + ScriptEngineMajorVersion()+"."+ScriptEngineMinorVersion()+"."+ScriptEngineBuildVersion() +"</b></td></tr>";
ret+= "<tr><td colspan=5>&nbsp;</td></tr>";

if (conStatus!="Outlook ActiveX")
{ret+= "<tr><td  class=printChars>&nbsp;</td><td class=printChars colspan=5><b></b> <a class=jellolinktop onclick=javascript:uninst();><font color=blue>" +txtUninst +"</a></td></tr>";}
ret+="<tr><td class=printChars>Last check for updates:</td><td class=printChars colspan=5>"+DisplayAppDate(new Date(jello.lastWebVersionCheck),false)+" <a class=jellolinktop onclick='checkVersion(true);'><b>Check now</b></a></td></tr>";

if (webJDversion!=jelloVersion && notEmpty(webJDversion)){ret+="<tr><td class=printChars colspan=5><font size=3 color=red>A new Jello Dashboard version is available. "+webJDversion+" </font><a href='http://www.jello-dashboard.net/download/' target='_blank'><b>Go to the downloads page</b></a></td></tr>";}
ret+="</table>";
return ret;
}


function langList()
{
//list of languages for settings
var ret="en;English;fr;French;de;German;es;Spanish;ru;Russian";
return ret;
}

function comboValue(a,b){
this.value=a;
this.text=b;
}


function updateOLCatAll()
{
//Update All+existing outlook categories from jello contexts/projects (OL2007 only)
var uCount=0;
var ds=tagStore;

var doOLCUpdate=function(r)
			{
			var itt=getTagType(r.get("id"));
			var ret=false;
			if (itt==true){ret=updateOLCategory(r.get("tag"));}
			if (ret==true){uCount++;}
			};

ds.each(doOLCUpdate);

if (uCount>0){alert(txtUpdateAllCatsInfo.replace("%1",uCount));}
}


function setJelloFolder(el)
{
var ss="";
var t=NSpace.PickFolder();
	if (t!=null)
	{

  if (el.getId()=="actionFolder" && t.DefaultItemType==3)
  {
  		el.setValue(t.FolderPath);
		var vel=Ext.getCmp("val"+el.getId());
		vel.setValue(t.EntryID);
		return t.EntryID;

  }

  if (el.getId()=="contactFolder" && t.DefaultItemType==2)
  {
  	el.setValue(t.FolderPath);
		var vel=Ext.getCmp("val"+el.getId());
		vel.setValue(t.EntryID);
		return t.EntryID;

  }

  if (el.getId()=="calendarFolder" && t.DefaultItemType==1)
  {
  	el.setValue(t.FolderPath);
		var vel=Ext.getCmp("val"+el.getId());
		vel.setValue(t.EntryID);
		return t.EntryID;

  }


    if (el.getId()=="nonActionFolder" && t.DefaultItemType==4)
  {
  		el.setValue(t.FolderPath);
		var vel=Ext.getCmp("val"+el.getId());
		vel.setValue(t.EntryID);
		return t.EntryID;

  }


  if (el.getId()=="actionFolder" && t.DefaultItemType!=3){alert(txtPromptTaskFolder);return false;}
  if (el.getId()=="contactFolder" && t.DefaultItemType!=2){alert(txtPromptContactFolder);return false;}
  if (el.getId()=="calendarFolder" && t.DefaultItemType!=1){alert(txtPromptCalendarFolder);return false;}
  if (el.getId()=="nonActionFolder" && t.DefaultItemType!=4){alert(txtPromptJournalFolder);return false;}

		if (t.DefaultItemType==0 || t.DefaultItemType==6)
		{
		el.setValue(t.FolderPath);
		var vel=Ext.getCmp("val"+el.getId());
		vel.setValue(t.EntryID);
		return t.EntryID;
		}
		else
		{alert(txtPromptMailFolder);return false;}
	}

}

function fchTriggerClick(el){
//handle folder selector trigger click
var fs=setJelloFolder(el);
if (fs==false || notEmpty(fs)==false){return;}
changeJelloSetting(el,fs);
}


//new
function backupSettings()
{
//get a backup of user's settings
var bup=jese.systemJournalEntry.Copy();

bup.Subject+=" "+txtBackup+"@"+new Date().format('j F Y (H:i)');
bup.Save();
Ext.info.msg(txtMsgBupTtl,txtMsgBup+":<br><a class=jellolink style=text-decoration:underline; onclick=olItem('"+bup.EntryID()+"');>"+bup+"</a>");
}

//new
function resetSettings()
{
//reset user's settings
var choice=confirm(txtMsgSetReset);
if (!choice){return;}
var bup=jese.systemJournalEntry.Copy();
bup.Subject+=" "+txtBackup+"@"+new Date().format('j F Y');
bup.Save();
jese.restoreDefaults();
Ext.info.msg(txtMsgResTtl,txtMsgSetReDone+":<br><a class=jellolink style=text-decoration:underline; onclick=olItem('"+bup.EntryID()+"');>"+bup+"</a>");
checkForReload("appLanguage");
}

//new
function sendSettings()
{
//send user's settings by mail
var newmsg=inboxItems.Items.Add();
newmsg.Attachments.Add(jese.systemJournalEntry);
newmsg.Display();
}

function importAllOLCategories(choose)
{
//get all outlook categories of action folder and create tags
if (choose)
{
//get from any folder
var jiF=NSpace.PickFolder();
if (jiF==null){return;}
var iF=jiF.Items;
var dasl=getInboxTaskDASL();
dasl=dasl.replace(" AND (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2)","");
//dasl=dasl.replace("OR (urn:schemas-microsoft-com:office:office#Keywords IS NULL)","OR (NOT urn:schemas-microsoft-com:office:office#Keywords IS NULL)");
}
else
{
//get from tasks
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var dasl=getInboxTaskDASL();
dasl=dasl.replace(" 2)"," 22)");
}


var tCounter=0;
var iits=iF.Restrict("@SQL="+dasl);

try{
	if (iits.Count>0)
	{
		for (var x=1;x<=iits.Count;x++)
		{
		var cats=iits(x).itemProperties.item(catProperty).Value;
		cats=cats.replace(new RegExp(",","g"),";");
			if (notEmpty(cats))
			{
				var cat=cats.split(";");
					if (cat.length>0)
					{
						for (var y=0;y<cat.length;y++)
						{
							var gtg=Trim(cat[y]);
							var ttg=getTag(gtg);
							if (ttg==null || typeof(ttg)=="undefined")
							{
							var tid=createTag(gtg,false,0,false,true);
							if (tid!=false)
								{tCounter++;}
								else
								{alert(gtg+" could not be added as tag");}
							}
						}
					}
			}
		}
	}
	else
	{
  //Ext.info.msg(txtImpOLCatTtl,'0 '+txtImpOlCasTgs);
  }
Ext.info.msg(txtImpOLCatTtl,'<b>'+tCounter+'</b> '+txtImpOlCasTgs);
}catch(e){alert(txtInvalid);}
}

function importNSpaceOLcategories()
{

	try{var tCounter=0;
	var cats = NSpace.Categories;
	for( var i=1; i <= cats.Count; i++){
		var cat = cats.Item(i);
		var cname = cat.Name;
		var gtg=Trim(cname);
		var ttg=getTag(gtg);
		if (ttg==null || typeof(ttg)=="undefined")
		{
		var tid=createTag(gtg,false,0,false,true);
		if (tid!=false)
			{tCounter++;}
			else
			{alert(gtg+" could not be added as tag");}
		}

	}
	Ext.info.msg(txtImpOLCatTtl,'<b>'+tCounter+'</b> '+txtImpOlCasTgs);
}catch(e){alert(txtInvalid);}
}

function analyzeObject(p)
{
var ret="";
  for (var x=0;x<p.length;x++)
  {
     var vp=p[x];
     for (prp in vp)
     {
     ret+=prp+":"+vp[prp]+"\n";
     }
  ret+="\n----\n";
  }

return ret;
}


function restoreSettings(w)
{
//Restore a settings set for use (after backup original)
var sval=Ext.getCmp("setlist").getValue();
var choice=confirm(txtSetResConfrm+"\n"+sval);
if (!choice){return;}
var iF=jese.systemJournalEntry.Parent.Items;
var j=iF.Restrict("@SQL=urn:schemas:httpmail:subject = '"+sval+"'");
	if (j.Count>0)
	{
		var it=j(1); //settings to be restored
		var itn=it.Copy();
		var cit=jese.systemJournalEntry; //current settings
		var ssname=cit.Subject;
		cit.Subject+=" "+txtSetBefRestore+" @ "+new Date().format('j F Y (H:i)');
		cit.Save();
		itn.Subject=ssname;
		itn.Save();
		w.destroy();
		Ext.info.msg(txtSetChngTtl,txtSetResOK);
		checkForReload("appLanguage");
	}

}

function openSettingsFolder()
{//open the settings folder
var jF=jese.systemJournalEntry.Parent;
var exp=jF.GetExplorer();exp.Display();
}

function fixOrphanTags()
{
 var cStore=tagStore;
 var orphans=new Array();
 var okCount=0;

	var checkTag=function(rec){
		var parent=rec.get("parent");
		var pt=getTagName(parent);
		if (pt==null){orphans.push(rec);}else{okCount++;}
	};

	cStore.each(checkTag);

	var oco=orphans.length;
		if (oco>0)
		{

			var newTag=0;
			//set all orphans to parent 0
			for (var x=0;x<oco;x++)
			{
				var r=orphans[x];
				r.beginEdit();r.set("parent",newTag);r.endEdit();tagStore.clearFilter();syncStore(tagStore,"jello.tags");
			}

		//try to restore root folders
		/*
		var ix=cStore.find("id",new RegExp("^1$"));
		var r=cStore.getAt(ix);if (r!=null){r.beginEdit();r.set("parent",0);r.endEdit();}
		var ix=cStore.find("id",new RegExp("^5$"));
		var r=cStore.getAt(ix);if (r!=null){r.beginEdit();r.set("parent",0);r.endEdit();}
		var ix=cStore.find("id",new RegExp("^10$"));
		var r=cStore.getAt(ix);if (r!=null){r.beginEdit();r.set("parent",0);r.endEdit();}
		tagStore.clearFilter();
		syncStore(tagStore,"jello.tags");
		*/

		jese.saveCurrent();
		alert(oco+" "+txtSIFixedOrphans);
		}
		else{alert(txtSINoFixedOrphans.replace("%1",okCount));}

}

//-----------------------
function viewSettings()
{
//view settings and restore selected set

//create the setting sets list
var settingsList=new Array();
var iF=jese.systemJournalEntry.Parent.Items;
var j=iF.Restrict("@SQL=urn:schemas:httpmail:subject LIKE '"+jelloSettingsName+"%'");
	if (j.Count>0)
	{
		for (var x=1;x<=j.Count;x++)
		{
    var sst=Ext.util.Format.htmlEncode(j(x));
    settingsList.push(sst);
    }
	}
var cset=Ext.util.Format.htmlEncode(jelloSettingsName);

var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: txtSetViewSets,
        height:370,
        //bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'settform',
        //iconCls:'tagformicon',
        buttonAlign:'center',
        defaults: {width: 480},
        defaultType: 'textfield',

        items: [
				new Ext.form.ComboBox({
                fieldLabel: txtSetViewSet,
                id: 'setlist',
                store:settingsList,
                hideTrigger:false,
                typeAhead:true,
                editable:false,
                forceSelection:true,
                triggerAction:'all',
                emptyText:txtSetSelSet,
                listeners:{select: function(cb,rec,idx){viewSettingsSet();}},
                mode:'local'
            }),
		        new Ext.form.Label({
          //      fieldLabel: txtSetVwConts,
                id: 'setview',
                style:'font-size:11px;font-family:Tahoma;overflow-y:scroll;overflow-x:scroll;width:880px;',
                height:220,
                width:800,
                cls:'settingslist',
                value:''
            })
        ]

    });



   var win = new Ext.Window({
        title: txtSettings,
        width: 680,
        height:390,
        id:'thesettform',
        minWidth: 300,
        minHeight: 200,
        resizable:false,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: simple,
        modal:true,
        listeners: {destroy:function(t){try{showOLViewCtl(true);}catch(e){}}},
        buttons: [{
            text:'<b>'+txtUpdate+'</b>',
            id:'usbut',
            tooltip:txtUpdateSettings,
            hidden:true,
            listeners:{
            click: function(b,e){
            updateChangedSettings(win);
            }
            }
        },{
            text:'<b>'+txtRestore+'</b>',
            id:'sbut',
            tooltip:txtSetRestIfo,
            listeners:{
            click: function(b,e){
            restoreSettings(win);
            }
            }
        },{
            text: txtCancel,
            tooltip:txtCancel,
                        listeners:{
            click: function(b,e){

            try{win.destroy();}catch(e){}
            }
            }
        }]
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);

    setTimeout(function(){
    Ext.getCmp("setlist").setValue(jelloSettingsName);
    viewSettingsSet();
    },3);
}

//new
function viewSettingsSet()
{
//update view to show contents of settings set
var sval=Ext.getCmp("setlist").getValue();
var sval=Ext.util.Format.htmlDecode(sval);
Ext.getCmp("setlist").setValue(sval);
var editMode=false;
if (sval==jelloSettingsName){editMode=true;}
var editLink="";
var lid=1;
var iF=jese.systemJournalEntry.Parent.Items;
var j=iF.Restrict("@SQL=urn:schemas:httpmail:subject = '"+sval+"'");
try{
	if (j.Count>0)
	{
	var stxt=j(1).UserProperties.item(1).Value;
	//Ext.getCmp("setview").setValue(stxt);

	var jjj=Ext.util.JSON.decode(stxt);var ret="";
	    for (prop in jjj)
    {
    var pp=jjj[prop];
    if (typeof(pp)=="object")
    {ret+="<img src=img//box.gif style=float:left;>&nbsp;<span class=jellolinkTopGlow style=color:red;><b><i>"+prop+"</i> [object]</b></span><br><br>"+analyzeObject(pp,editMode,prop,lid)+"<br>";}
    else
		{
		if (editMode){editLink="<a id='setlink"+lid+"' class=jellolinktop style=text-decoration:underline; onclick=editsetng("+lid+",'"+prop+"');>";}
		ret+="<img src=img//info.gif style=float:left;>&nbsp;<span><b>"+prop+"</b></span><br>"+editLink+pp+"</a><hr><br>";
		}
	lid++;
    }
	Ext.getCmp("setview").getEl().update(ret);
  }
  }catch(e){Ext.getCmp("setview").getEl().update("<img src=img//icon_alert.gif><br><h1>"+txtSetCannotView+"</h1>");}
}

function analyzeObject(p,editMode,prop,lid)
{
lid=100*lid;
var ret="";var editLink="";
  for (var x=0;x<p.length;x++)
  {
     var vp=p[x];
     var idx=0;
     for (prp in vp)
     {
     if (editMode){
	 var pp="";
     editLink="<a id='setlink"+lid+"' class=jellolinktop style=text-decoration:underline; onclick=editsetng("+lid+",'"+prop+"',"+x+",'"+prp+"');>";}
     pp=vp[prp];
     ret+="<div class=notag><b>"+prp+"</b> : "+editLink+pp+"</a><br></div>";
     idx++;
     lid++;
     }
  ret+="<br>";
  }

return ret;
}

//new
function editsetng(lid,prop,rec,fld)
{
//edit setting value
var defval="";
if (typeof(rec)=="undefined" && typeof(fld)=="undefined")
	{//simple value
	eval("defval=jello."+prop);
	}
	else
	{//object value
	eval("var obval=jello."+prop+"["+rec+"]");
	defval=obval[fld];
	prop=prop+"["+rec+"]."+fld;

	}
var pv=prompt("Edit Settings Value for ["+prop+"]",defval);

if (!notEmpty(pv)){return;}
	try{
	eval("jello."+prop+"="+pv);
	}catch(e){
	eval("jello."+prop+"='"+pv+"'");
	}

var ae=document.getElementById("setlink"+lid);
ae.innerHTML="<span class=tagprj>"+pv+"</span>";
Ext.getCmp("usbut").show();
}

//new
function updateChangedSettings(win)
{
var uss=confirm(txtUpdateCustSetAsk);
if (uss)
{jese.saveCurrent();
checkForReload("appLanguage");
Ext.getCmp("thesettform").destroy();
}
}

function updateUseOfMail(vl)
{
	if (vl==1)
	{
	Ext.getCmp("inboxFolder").disable();
	Ext.getCmp("archiveFolder").disable();
	//Ext.getCmp("autoUpdateSecuredFields").disable();
	Ext.getCmp("mailPreviewHTML").disable();
	Ext.getCmp("inboxAutoRefresh").disable();
	Ext.getCmp("markReadfromJello").disable();
	Ext.getCmp("useGoToFolderToMove").disable();
	}
	else
	{
	Ext.getCmp("inboxFolder").enable();
	Ext.getCmp("archiveFolder").enable();
	//Ext.getCmp("autoUpdateSecuredFields").enable();
	Ext.getCmp("mailPreviewHTML").enable();
	Ext.getCmp("inboxAutoRefresh").enable();
	Ext.getCmp("markReadfromJello").enable();
	Ext.getCmp("useGoToFolderToMove").enable();
	}
}



