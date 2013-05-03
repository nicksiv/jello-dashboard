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
function loadMigrationScr(){}

function settingsMigration(check)
{
//check for existing previous version settings to migrate
  var je=journalItems.Items;
  var ret="";
	var j=je.Restrict("[Subject]='<Jello.Dash Settings>'");
  if (j.Count>0)
  {//migrate contexts and projects
  var set4=j(1);
  if (check==true){ret="<a class=jellolinktop onclick=settingsMigration()><font color=red>"+txtMigHomeMsg+"</font></a>";return ret;}
  var choice=confirm(txtMigPrmpt1);
  if (!choice){return;}
  var j4=new Array();
  var counter=0;
  //read all settings from v4
	for (var x=0;x<100;x++)
	{
	j4[x]=getCookie(set4,x);
	}
  
 var folds=	j4[36]+"|"+j4[37];
 var props=	j4[56];

 //create folders to host imported items
 var impTags=createTag(txtMigImpTags,true,0,true,true); 
 var arcTags=createTag(txtMigArcPrj,true,0,true,true); 
 
 var choice=confirm(txtMigPrmpt2);
 if (choice){migFols=migrateGetFolders(j4[69],(j4[6]=="1"));}
  
 //get all contexts and projects and create tags
 window.status=txtMigTgsCre;
 //contexts
  var j4cplist=j4[1];
  cpList=j4cplist.split("|");
    for (var x=0;x<cpList.length;x++)
    {
      var nt=cpList[x];
        if (notEmpty(nt))
        {
          var tnt=Trim(nt);
          var r=getTag(tnt);
          if (typeof(r)=="undefined")
          {
          var newTag=createTag(tnt,true,impTags,false,true);counter++;
          }else{newTag=r.get("id");}
          try{migrateCPProperties(tnt,newTag,folds,props);}catch(e){}
        }     
    }
 //projects
 window.status=txtStatusGetting;
  var j4cplist=j4[30]+"|"+j4[40];
  cpList=j4cplist.split("|");
    for (var x=0;x<cpList.length;x++)
    {
      var nt=cpList[x];
        if (notEmpty(nt))
        {
          var tnt=Trim(nt);
          var r=getTag(tnt);
			if (typeof(r)=="undefined")
			{
				if (tnt.substr(0,1)=="~")
				{tnt=tnt.replace(new RegExp("~","g"),"");
				var newTag=createTag(tnt,true,arcTags,false,true);counter++;}
				else
				{var newTag=createTag(tnt,true,10,false,true);counter++;}
			}else{newTag=r.get("id");}
			try{migrateCPProperties(tnt,newTag,folds,props);}catch(e){}
        }     
    }

//basic contexts names
var ix=tagStore.find("id",new RegExp("^6$"));var r=tagStore.getAt(ix);
	if (r!=null){r.beginEdit();r.set("tag",Trim(j4[2]));r.endEdit();tagStore.clearFilter();syncStore(tagStore,"jello.tags");}
var ix=tagStore.find("id",new RegExp("^8$"));var r=tagStore.getAt(ix);
	if (r!=null){r.beginEdit();r.set("tag",Trim(j4[63]));r.endEdit();tagStore.clearFilter();syncStore(tagStore,"jello.tags");}
var ix=tagStore.find("id",new RegExp("^9$"));var r=tagStore.getAt(ix);
	if (r!=null){r.beginEdit();r.set("tag",Trim(j4[64]));r.endEdit();tagStore.clearFilter();syncStore(tagStore,"jello.tags");}
var ix=tagStore.find("id",new RegExp("^7$"));var r=tagStore.getAt(ix);
	if (r!=null){r.beginEdit();r.set("tag",Trim(j4[65]));r.endEdit();tagStore.clearFilter();syncStore(tagStore,"jello.tags");}


//rest of settings
window.status=txtStMsgUpdating;
  jello.updateOutlookCategories=(j4[94]=="1");
  jello.timeFormat=j4[10];  
  jello.dateSeperator=j4[11];
  jello.agendaDays=j4[14];
  jello.dueDisplayDaysLeft=(j4[15]=="1");
  jello.appTitle=j4[16];
  jello.appLanguage=j4[44];
  jello.calendarShowCompleted=(j4[31]=="1");
  jello.sidebarLeft=(j4[41]=="1");
  jello.pageSize=j4[5];
  jello.inboxFolder=j4[67];
  jello.showNotStartedAlways=j4[71];
  jello.newActionStatus=j4[88];
  jello.dateFormat=j4[9];
  jello.defaultCalendarView=j4[92];
  
    alert(txtMigMsgSucc.replace("%1",counter));
    jello.migrated=true;
    jese.saveCurrent();
    migFols=null;
    window.status=txtReady;
    location.reload();
  }
  
  if (check==true){return ret;}
}

function getCookie(theJournalEntry,id)
{
//get a setting from jello 4 settings set
var gco=null;
	try{
	gco=theJournalEntry.UserProperties("js-"+id).Value;}
	catch(e){}
return gco;
}

function migrateGetFolders(SIncludeFolders,tasksOnly)
{
//enumerate all folders in default mailbox
folds=new Array();

var par=NSpace.GetDefaultFolder(6).Parent;
var folChild;

//add default folders
		if (tasksOnly==true)
		{
		folds=new Array();
		folds.push(taskItems);
		}
		else
		{
		folds=new Array();
		folds.push(taskItems);
		folds.push(calendarItems);
		folds.push(contactItems);
		folds.push(inboxItems);
		folds.push(noteItems);
		}
//add included folders
    if (SIncludeFolders!=null && SIncludeFolders!="" && SIncludeFolders.length>2)
    {
    var o=SIncludeFolders.split("|");
        for (var x=0;x<=o.length;x++)
        {
            try
            {
            if (o[x]!=null && o[x]!="" && o[x].length>2 && typeof(o[x])!="undefined")
            {var ff=NSpace.GetFolderFromID(o[x]);
            if (tasksOnly==true)
            {if (ff.DefaultItemType==3){folds.push(ff);}}
            else
		    {folds.push(ff);}
		    }
		    }catch(e){}
		}
    }
return folds;
}


function migrateCPProperties(tnt,newTag,folds,props)
{
//migrate contexts and project properties
var tarcFolder="";
var tprivate=false;
var tnotes="";
status=txtMsgProc+" "+tnt;
	var fs=folds.split("|");
	//check archive folder of tag
		if (fs.length>0)
		{
			for (var x=0;x<fs.length;x++)
			{
				var fss=fs[x].split(":");
					if (trimLow(tnt)==trimLow(fss[0]))
					{tarcFolder=Trim(fss[1]);}
			}
		}

	var fs=props.split("|");
	//check properties of tag
		if (fs.length>0)
		{
			for (var x=0;x<fs.length;x++)
			{
				var fss=fs[x].split("^");
					if (trimLow(tnt)==trimLow(fss[0]))
					{tnotes=Trim(fss[1]);
					var pv=fss[2];
					if (pv=="1"){tprivate=true;}
					}
			}
		}
var ix=tagStore.find("id",new RegExp("^"+newTag+"$"));
var r=tagStore.getAt(ix);
	if (r!=null)
	{
	r.beginEdit();
	r.set("notes",tnotes);
	r.set("private",tprivate);
	r.set("archive",tarcFolder);
	r.endEdit();
tagStore.clearFilter();
	syncStore(tagStore,"jello.tags");
	//save settings
	jese.saveCurrent();

//create tasks for each action item (and copy tasks if not in action folder)
//loop migFols
 if (migFols==null){return;}
      if (migFols.length>0)
      {
      for (var x=0;x<migFols.length;x++)
      {
          var f=null;
          try{f=setAndCheckArcFolder(migFols[x].EntryID);}catch(e){}
            if (f!=null)          
            {
            migrateScatteredItems(f,tnt);
            }
      }
      }
	}
}

function migrateScatteredItems(f,tagName)
{
//create tasks for all items existed in old version of Jello
if (f.EntryID==jello.actionFolder){return;}
var actionF=setAndCheckArcFolder(jello.actionFolder);
if (actionF==null){status=txtMigStMsgInvAF;return;}
var iF=f.Items;
var ctx=tagName;
var dasl="(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "') AND (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2)";
var its=iF.Restrict("@SQL="+dasl);
counter=its.Count;
if (counter==0){return;}
    for (var x=1;x<counter;x++)
    {
      var it=its(x);
      status=txtMsgProc+it;
        if (it.Class==48)
        {
        //task items. Just create a copy into current action folder
        try{
          var nit=it.Copy();
          var ni=nit.Move(actionF);
          }catch(e){}
        }
        else
        {
        try{var ait=archiveItem(it,null,null,true,false,true);}catch(e){}
        }
    }
}