//------------------- tag -------------
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
//    2008-2013 N.Sivridis http://jello-dashboard.net


//-------------------------------------

//---hta installer

Ext.BLANK_IMAGE_URL = 'img\\s.gif';
Ext.EventManager.on(window, 'load', function(){
    	setTimeout(function(){
        Ext.get('loading').remove();
        Ext.get('loading-mask').fadeOut({remove:true});
    }, 250);
});


var panelHeight=document.body.clientHeight-70;
document.title="Jello.Dashboard Control Center";

Ext.onReady(function(){
		Ext.QuickTips.init();
paintScreen();
//temp
//pTodotxt();

        });

var appName="jelloapp.hta";
var setIntoFolder="";

function validateAllFolders(){}
function updateCSS(){}
function pHome(){}

function paintScreen()
{

var rfun1="<a class=jellolink onclick=javascript:pQuickSetup();><img class='htaimg' src='img\\ud-tags (36).gif' />Quick Setup</a><br>";
rfun1+="<a class=jellolink onclick=javascript:pOutlookRun();><img class='htaimg' src='img\\ud-folder (14).gif' />Run in Outlook</a><br>";
rfun1+="<a class=jellolink onclick=javascript:pOutlookInfo();><img class='htaimg' src='img\\info.gif' />Outlook information</a>";
//rfun1+="<br><a class=jellolink href='#' onclick=javascript:pInstall();><img class='htaimg' src='img\\ud-tags (15).gif' /><b>Advanced Setup</b></a>";


var rfun3="<a class=jellolink onclick=pRun('jello5.hta')><img class='htaimg' src='img\\review.gif' />Run Jello without Outlook</a><br>";
rfun3+="<a class=jellolink onclick=pRun2('jello5.htm')><img class='htaimg' src='img\\page_url.gif' />Run in Internet Explorer</a><br>";
rfun3+="<a class=jellolink onclick=pRun('help.pdf')><img class='htaimg' src='img\\icon_info.gif' />"+txtHelp+"</a><br>";
var rfun4="<a class=jellolink onclick=openWebLinkinDefaultBrowser('http://jello-dashboard.net');><img class='htaimg' src='img\\j-icon.gif' />Jello Web page</a><br>";
rfun4+="<a class=jellolink onclick=openWebLinkinDefaultBrowser('http://www.jello-dashboard.net/forum');><img class='htaimg' src='img\\ud-tags (7).gif' />Jello Forums</a><br>";
rfun4+="<a class=jellolink onclick=openWebLinkinDefaultBrowser('http://sourceforge.net/projects/jello5/');><img class='htaimg' src='img\\ud-tags (48).gif' />Sourceforge page</a><br>";
rfun4+="<a class=jellolink onclick=openWebLinkinDefaultBrowser('https://github.com/nicksiv/jello-dashboard/');><img class='htaimg' src='img\\ud-tags (49).gif' />Github source code</a><br>";


var rfun2="<a class=jellolink onclick=javascript:pSettings();><img class='htaimg' src='img\\settings16.png' />"+txtInSetLoca+"</a>";
rfun2+="<br><a class=jellolink onclick=javascript:pImpExp();><img class='htaimg' src='img\\ud-tags (4).gif' />Settings Import/Export</a>";

var rfun5="<a class=jellolink onclick=javascript:pTodotxt();><img class='htaimg' src='img\\ud-folder (21).gif' />Export Todo.txt</a>";
rfun5+="<br><a class=jellolink onclick=javascript:pWiki();><img class='htaimg' src='img\\page_url.gif' />Export to Tiddly Wiki</a>";

var vp=new Ext.Viewport({
    layout: 'border',
    items: [ {
        region: 'west',
        collapsible: false,
        //title: 'Jello Dashboard',
        split:true,
	margins:'5 1 5 5',
	border:false,
	autoHeight:true,
	layout:'anchor',
	//layout:'accordion',
	minWidth: 200,
	maxWidth:400,
	width:200,
		items:[
			{title:'Outlook Setup',bodyStyle:'padding:6px',html:rfun1},
		    	{title:'Settings Storage',bodyStyle:'padding:6px',html:rfun2},
		    	{title:'Export data',bodyStyle:'padding:6px',html:rfun5},

		    	{title:'Run Jello',bodyStyle:'padding:6px',html:rfun3},
		    	{title:'Links',bodyStyle:'padding:6px',html:rfun4,collapsible:true,collapsed:true}
			]
    	},  {
        region: 'center',
        title: 'Title for the Grid Panel',
        collapsible: false,
        split: true,
	id:'mainpanel2',
	margins:'5 5 5 1',
	layout:'anchor',
        autoWidth:true
    } 
    ]
});
pQuickSetup();
}


function pInstall()
{

    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");
   var defFol=journalItems.Parent;

var simple = new Ext.FormPanel({
        labelWidth: 180,
        frame:true,
        title:txtInOLSetUp,
        renderTo:'mainpanel2',
    	height:250,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'tagform',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',

        items: [

            new Ext.form.TriggerField({
            	fieldLabel:txtInSelOLFol,
            	value:defFol.FolderPath,
            	width:250,
            	hidden:false,
            	onTriggerClick:function(e){selectInstFolder();},
            	id:'fname'
            	}),
            	new Ext.form.Hidden({
            	value:defFol.EntryID,
            	id:'fID'
              }),
              new Ext.form.Label({
              id:'finfo',
              width:400,
              html:''
              })
        ],
		buttons:[{
            text:'<b>'+txtSet+'</b>',
            id:'sins',
            disabled:true,
            tooltip:txtInSetInfo,
            listeners:{click: function(b,e){installAtFolder(1);}}
				},{
            text: txtUnset,
            disabled:true,
            id:'suni',
            tooltip:txtInUnSetInfo,
            listeners:{click: function(b,e){installAtFolder(0);}}
            }]
        
		
    });
    
simple.render(mp.getEl());
//mp.doLayout();
finfo.innerHTML=updateFolderInfo(defFol.EntryID);
    
}

function updateFolderInfo(fid)
{
if (fid==null){fid=Ext.getCmp("fID").getValue();}
var ff=NSpace.GetFolderFromID(fid);
var ret="<div style=font-size:12px;width:430px;padding-left:180px;padding-top:30px;>";
var ib=Ext.getCmp("sins");
var ub=Ext.getCmp("suni");
var hasHomepage=false;
var homepage="";
var jelloOn=false;
hasHomepage=ff.WebViewOn;
homepage=ff.WebViewURL;
if (hasHomepage){ret+="<b>"+txtInWVIsOn+"</b><br><br>";}
else{if (!notEmpty(homepage)){ret+=txtInWVIsOff+"<br><br>";}else{ret+="<b>"+txtInWPOn+"</b> "+txtInWpOnWvOff+". <a class=jellolinktop onclick=setWViewOn('"+fid+"');><b>"+txtInWpClOn+"</b></a><br><br>";}}
if (notEmpty(homepage)){ret+="<a class=jellolinktop href='"+homepage+"'>"+homepage+"</a><br>";}
if (homepage.search("jello5.htm")>-1 && hasHomepage){ret+="<br><span class=tagprj>"+txtInJ5On+"</span>";jelloOn=true;}
if (homepage.search("jelloDash.htm")>-1 && hasHomepage){ret+="<br><span class=tagcon>"+txtInJ4On+"</span>";jelloOn=true;}
if (jelloOn){ib.disable();ub.enable();}else{ub.disable();ib.enable();ret+="<br><br>"+txtInNoJello;}
ret+="</div>";
return ret;
}

function selectInstFolder()
{
var ff=NSpace.PickFolder();
if (!notEmpty(ff)){return;}
Ext.getCmp("fID").setValue(ff.EntryID);
Ext.getCmp("fname").setValue(ff.FolderPath);
finfo.innerHTML=updateFolderInfo(ff.EntryID);
}

function setWViewOn(fid)
{
var ff=NSpace.GetFolderFromID(fid);
ff.WebViewOn=true;
finfo.innerHTML=updateFolderInfo(fid);
}

function installAtFolder(fun)
{
var fid=Ext.getCmp("fID").getValue();
var ff=NSpace.GetFolderFromID(fid);

	if (fun==0)
	{
	//uninstall
		
	ff.WebViewOn=false;
	finfo.innerHTML=updateFolderInfo(fid);	
	alert(txtInJInstOK.replace("%1",ff));
	}
	else
	{
	//install
	var uri=getAppPath();
	uri=uri.replace(appName,"jello5.htm");
	uri=uri.replace(new RegExp("\/","g"),"\\");
	ff.WebViewOn=true;	
		try{
		ff.WebViewURL=uri;
		}catch(e){alert(txtInInvalidSetFldr);}
		if (ff.WebViewURL!=uri){alert(txtInInvalidSetFldr);}
	}
	
finfo.innerHTML=updateFolderInfo(fid);
}

function pSettings()
{
 var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");
	var ff=jelloSettingsFolder;
	if (!notEmpty(ff)){ff=journalItems.EntryID;}
	var sfol=NSpace.GetFolderFromID(ff);
	
var simple = new Ext.FormPanel({
        labelWidth: 180,
        frame:true,
        title:txtInSetLoca,
        height:350,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'tagform',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',

        items: [
			 {
                fieldLabel: txtInSetEnNm,
                id: 'fname',
                value:jelloSettingsName,
                disabled:false,
                tabIndex:0,
                allowBlank:false
            },
            new Ext.form.TriggerField({
            	fieldLabel:txtJSetFol,
            	value:sfol.FolderPath,
            	width:300,
            	hidden:false,
            	style:'margin-left:0px',
            	onTriggerClick:function(e){selectJouFolder();},
            	id:'ffol'
            	}),    
              new Ext.form.Checkbox({
            	fieldLabel:txtInSetMove,
            	id:'fmove',
            	checked:true
            	
           		}),
            	new Ext.form.Hidden({
            	value:sfol.EntryID,
            	id:'fID'
              }),
              new Ext.form.Label({
              id:'finfo',
              width:400,
              html:'<div style=font-size:12px;width:430px;padding-left:180px;padding-top:30px;>'+txtInJouSetInfo+'</div>'
              })
        ],
		buttons:[{
            text:'<b>'+txtUpdate+'</b>',
            id:'supd',
            tooltip:txtUpdate,
            listeners:{click: function(b,e){changeJournalLocation();}}
				},{
            text:txtSetDefaults,
            id:'sdefs',
            tooltip:txtSetDefaults,
            listeners:{click: function(b,e){setJournalDefs();}}
				},{
            text:txtBackup,
            id:'sbup',
            tooltip:txtSIBackup,
            listeners:{click: function(b,e){backupSettingsInst();}}
				}]
        
		
    });
    
simple.render(mp.getEl());
 
}

function selectJouFolder()
{
var ff=NSpace.PickFolder();
if (!notEmpty(ff)){return;}
if (ff.DefaultItemType!=4){alert(txtJouFolderPrompt);return;}
Ext.getCmp("fID").setValue(ff.EntryID);
Ext.getCmp("ffol").setValue(ff.FolderPath);
}

function changeJournalLocation()
{
var uri=getAppPath();
uri=uri.replace(appName,"");
var ff=jelloSettingsFolder;
if (!notEmpty(ff)){ff=journalItems.EntryID;}
var newName=Ext.getCmp("fname").getValue();
var newFol=Ext.getCmp("fID").getValue();
try{var theNewFolder=NSpace.GetFolderFromID(newFol);}catch(e){alert(txtInvalid);return;}
var nameChange=false;
var folderChange=false;
if (newName!=jelloSettingsName){nameChange=true;}
if (newFol!=ff){folderChange=true;}
if (newFol==journalItems.EntryID){newFol="";}
var fso=new ActiveXObject("Scripting.FileSystemObject");

var furi=uri+"/js/customSettings.js";

fl = fso.CreateTextFile(furi, 1);
fl.WriteLine("var jelloSettingsName='"+newName+"';");
fl.WriteLine("var jelloSettingsFolder='"+newFol+"';");
fl.Close();
var xtr="";
var toMove=Ext.getCmp("fmove").getValue();
if (toMove)
  { 
  if (nameChange){jese.systemJournalEntry.Subject=newName;jese.systemJournalEntry.Save();}
	if (folderChange)
	{var nj=jese.systemJournalEntry.Copy();
	nj.Move(theNewFolder);
	xtr=txtInSettngMovd;
	}
	}
alert(txtInSetChgMsg+"\n"+xtr);
}

function setJournalDefs()
{
Ext.getCmp("fname").setValue("<Jello.5 Settings>");
var df=journalItems;
Ext.getCmp("fID").setValue(df.EntryID);
Ext.getCmp("ffol").setValue(df.FolderPath);
alert(txtInSetDefsMsg);
}



function pRun(fl)
{

var uri=getAppPath();
uri=uri.replace(appName,"");
uri=uri.replace("#","");
var furi=uri+fl;
furi='"'+furi+'"';
alert(furi);
var fso=new ActiveXObject("WScript.Shell");
fso.run(furi);
}


function pRun2(fl)
{
var uri=getAppPath();
uri=uri.replace(appName,"");
var furi=uri+fl;
furi='iexplore "'+furi+'"';
var fso=new ActiveXObject("WScript.Shell");
fso.run(furi);
}



function backupSettingsInst()
{
//get a backup of user's settings
var bup=jese.systemJournalEntry.Copy();
bup.Subject+=" "+txtBackup+"@"+new Date().format('j F Y (H:i)');
bup.Save();
alert(txtMsgBup+" : "+bup);
}

function pImpExp()
{
//import/export
var defef=jello.setExportFile;
var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");

  if (notEmpty(defef))
  {
  var furi=defef
  }
  else
  {
  var uri=getAppPath();
  var furi=uri+"/jelloSettings.txt";
  furi=furi.replace("\/"+appName,"");
  furi=furi.replace("#","");
  }
ret="<p align=center><br><br><img src=img//biglogo.gif><br></p><div style='position:absolute;left:30px;align=center'><br><br><h2>Export Settings to file</h2>";
ret+="<input class=htmlControls id=expfiled type=text size=80 value='"+furi+"'>&nbsp;"
ret+="<input class=htmlControls type=button onclick=exportFile() value='Export Settings'><br></p><br><br><br>";
ret+="<hr><br><h2>Import Settings from file</h2>";
ret+="<input type=file class=htmlControls id=impfiled value='"+furi+"' size=80 value='Import Settings'>&nbsp;";
ret+="<input type=button class=htmlControls onclick=importFile() value='Import Settings'><br><br>";
ret+="</div></p>";
mp.getEl().update(ret);

}

function exportFile()
{
//export jello settings to a text file
var fn=document.getElementById("expfiled").value;
try{    
var fso=new ActiveXObject("Scripting.FileSystemObject");
var t=new Date();
var jsonStr=Ext.util.JSON.encode(jello);
fl = fso.CreateTextFile(fn, 1);
fl.WriteLine("[JELLO.DASHBOARD SETTINGS FILE] "+t.toLocaleDateString()+ " "+DisplayTime(t)+"\n");
fl.WriteLine(jsonStr);
fl.Close();
fso=null;
jello.setExportFile=fn;
jese.saveCurrent();
alert("Settings exported"); 
}catch(e){alert("An Error Occured.\n File name invalid or Filesystem object is blocked from security!");}
}

function importFile()
{
//import jello settings from a text file
var fn=document.getElementById("impfiled").value;

//try{    
var fso=new ActiveXObject("Scripting.FileSystemObject");


//var jsonStr=Ext.util.JSON.decode(jello);
fl = fso.OpenTextFile(fn, 1);
var ttl=fl.ReadLine();
var contt=fl.Read(99999);
fl.Close();
  if(ttl.substr(0,31)=="[JELLO.DASHBOARD SETTINGS FILE]")
  {
  fso=null;
  var choice=confirm("Do you want to import thie settings file?\n(a journal entry backup will be made) \n\n"+ttl);
  if (choice)
  {
  backupSettings();
  jello=Ext.util.JSON.decode(contt);
  jese.saveCurrent();
  alert("Settings imported successfuly");
  } 
  }else{alert("Invalid file!");}
//}catch(e){alert("An Error Occured.\n File name invalid or Filesystem object is blocked from security!");}

}

function backupSettings()
{
//get a backup of user's settings
var bup=jese.systemJournalEntry.Copy();

bup.Subject+=" "+txtBackup+"@"+new Date().format('j F Y (H:i)');
bup.Save();
}

function openWebLinkinDefaultBrowser(url)
{
      var shell = new ActiveXObject("WScript.Shell");
      shell.run(url);
}

function pQuickSetup()
{
//quick and automated setup in outlook

    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");

    var msg="<table width=80% align=center><tr><td width=300 height=100 align=middle bgcolor=lightgray><a class=jellolinktop onclick=quick1();>Setup to Outlook</a>"+isJelloInstalled()+"</td><td width=20>&nbsp;</td><td align=middle width=300 bgcolor=pink><a class=jellolinktop onclick=quick2();>Remove from Outlook</a></td></tr></table></p><br><br><br>";


var simple = new Ext.FormPanel({
        labelWidth: 180,
        frame:false,
        title:'Outlook Setup',
        renderTo:'mainpanel2',
    	autoHeight:false,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'quickform',
        buttonAlign:'center',
    	html:'<br><br><p align=center><img src=img\\biglogo.gif><br><br><b>Welcome to Jello Dashboard Quick Setup</b><br>'+txtVersion+': <i><font color=green>'+jelloVersion+'</font></i><br><br><br><br>'+msg
    });


    var mp=Ext.getCmp("mainpanel2");

simple.render(mp.getEl());


}

function getAllOLFolders()
{
//get all outlook folders

    var flds=new Array();

    for (var x=1;x<=NSpace.Folders.Count;x++)
	{
	   try{ flds.push(NSpace.Folders(x));}catch(e){}
	}
    var y=0;
    do
    {

	for (var z=1;z<=flds[y].Folders.count;z++)
	{
	    try{flds.push(flds[y].folders(z));}catch(e){}
	    x++;
	}
	y++;
    } while (x<flds.length)


return flds;

}

function getJelloInstFolder()
{
//get all outlook folder with jello installation info
var ret="<b>Select one or more folders where <font color=green>Jello Dashboard</font> page should be linked</b><br>Please note that setting a folder's view to Jello will not harm the specific folder's data and can be easily removed.<hr><table width=95% class=showtable>";

 	var uri=getAppPath();
    	uri=uri.replace(appName,"jello5.htm");
    	uri=uri.replace(new RegExp("\/","g"),"\\");

var olFolders=getAllOLFolders();
    for (var x=0;x<olFolders.length;x++)
	{
	    var fd=olFolders[x];
	    var hasHomepage=fd.WebViewOn;
	    var homepage=fd.WebViewURL;

	    if (homepage.search("jello5.htm")>-1 && hasHomepage)
		{ret+="<tr><td width=40% height=30><img src=img\\database.gif>&nbsp;<b>"+(x+1)+". "+fd + "</b></td><td><img src=img\\j-icon.gif>&nbsp;<span class=tagcon>Jello Dashboard</span></td><td><a class=jelloformlink onclick=removeJelloHere('"+fd.EntryID+"')>Remove</a></td></tr>";}
	    else
		{
		    
		try{
		    var noFlag=0;
		    var olduri=fd.WebViewURL;
		    var oldwv=fd.WebViewOn;
		    fd.WebViewURL=uri;
		    fd.WebViewOn=true;
		    if (fd.WebViewURL!=uri){ret+="<tr><td width=40% height=30><img src=img\\delete.gif>&nbsp;<i>"+(x+1)+". "+fd+"</i></td><td><span class=tagprj>Unavailable</span></td><td></td></tr>";noFlag=1;}
		    fd.WebViewURL=olduri;
		    fd.WebViewOn=oldwv;
		    if (noFlag==0){ret+="<tr><td width=40% height=30><img src=img//folder.gif>"+(x+1)+". "+fd.FolderPath+"</td><td></td><td><a class=jelloformlink onclick=installJelloHere('"+fd.EntryID+"')>Install</a></td></tr>";}
		   }catch(e){ret+="<tr><td width=40% height=30><img src=img\\delete.gif>&nbsp;<i>"+(x+1)+". "+fd+"</i></td><td></td><td><span class=tagprj>Unavailable</span></td></tr>";}
   
		}
	    
	}
return ret+"</table>";
}


function quick1()
{
//quick setup jello into an Outlook folder
    var onFolder=getJelloInstFolder();

    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");


var simple = new Ext.FormPanel({
        labelWidth: 180,
        frame:false,
        title:'Setup to Outlook',
        renderTo:'mainpanel2',
    	autoHeight:false,
        height:450,
    	autoScroll:true,
    	layout:'fit',
    	bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'quickform',
        buttonAlign:'center',
    	html:onFolder
    });


    var mp=Ext.getCmp("mainpanel2");

simple.render(mp.getEl());


}

function installJelloHere(fid)
{
//install jello to a folder
var jdf=NSpace.GetFolderFromID(fid);
var choice=confirm("Do you want to setup Jello in "+jdf.FolderPath+ "?");
  
if (choice)
    { 	var uri=getAppPath();
    	uri=uri.replace(appName,"jello5.htm");
    	uri=uri.replace(new RegExp("\/","g"),"\\");
	jdf.WebViewURL=uri;
	jdf.WebViewOn=true;
	quick1();
    }

}

function removeJelloHere(fid)
{
//install jello to a folder
var jdf=NSpace.GetFolderFromID(fid);
var choice=confirm("Do you want to remove Jello from "+jdf.FolderPath+ "?");
  
if (choice)
    { 
    
	jdf.WebViewURL="";
	jdf.WebViewOn=false;
	quick1();
    }

}

function quick2()
{
//remove jello from all folders

  var choice=confirm("Do you want to remove Jello Dashboard links from your Outlook folders?\nYou can set it up again any time!");
    if (!choice){return;}
    var olFolders=getAllOLFolders()
    var found=0;
    for (var x=0;x<olFolders.length;x++)
	{
	    var fd=olFolders[x];
	    var hasHomepage=fd.WebViewOn;
	    var homepage=fd.WebViewURL;

	    if (homepage.search("jello5.htm")>-1 && hasHomepage)
		{
		    fd.WebViewURL="";
		    fd.WebViewOn=false;
		    found++;
		}
	}
    alert(found+" Outlook folders updated\nJello Dashboard was successfully unlinked from your Outlook folders");
    pQuickSetup();
}

function isJelloInstalled()
{
    var olFolders=getAllOLFolders()
    var found="<br>&nbsp;";
    var fcount=0;
    for (var x=0;x<olFolders.length;x++)
	{
	    var fd=olFolders[x];
	    var hasHomepage=fd.WebViewOn;
	    var homepage=fd.WebViewURL;

	    if (homepage.search("jello5.htm")>-1 && hasHomepage)
		{
		   // found+="<span class=showtable>"+fd + " <b>|</b></span>&nbsp;";
		    setIntoFolder=fd.FolderPath;
		    fcount++;
		}
	}
if (fcount==0){found="";}
return found+"<span class=showtable>Currently linked to "+fcount+" Outlook folder(s)</span>";
}

function pOutlookRun()
{
//run outlook and display jello folder
var furi="outlook.exe /select outlook:"+setIntoFolder
var fso=new ActiveXObject("WScript.Shell");
fso.run(furi);

}

function pOutlookInfo()
{
//display general info
    var gset=genSettings();
    gset=gset.replace("\/"+appName,"");
    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");


var simple = new Ext.FormPanel({
        labelWidth: 180,
        frame:false,
        title:'Outlook information',
        renderTo:'mainpanel2',
    	autoHeight:false,
        height:450,
    	autoScroll:true,
    	layout:'fit',
    	bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'infoform',
        buttonAlign:'center',
    	html:gset
    });



simple.render(mp.getEl());


}


function htaResize()
{
	if (conStatus=="Outlook ActiveX")
{
if(typeof(oJello.applicationName)!="undefined")
	{
	jello.htaWidth=document.body.clientWidth;
	jello.htaHeight=document.body.clientHeight;
	jese.saveCurrent();
	}
}
}

function pWiki()
{
//export to a tiddlywiki
// tiddly.js
tiddlyScreen();
}

function pTodotxt()
{
//sync with Todo.txt file
todoTxtSync();
}
