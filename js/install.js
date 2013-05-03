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
//    2008-2012 N.Sivridis http://jello-dashboard.net


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
        });

var appName="install.hta";

function validateAllFolders(){}
function updateCSS(){}
function pHome(){}

function paintScreen()
{

new Ext.Panel({
                    border:true,
                    renderTo:'main',
                    id:'centerpanel',
                    border:false,
                    deferredRender:false,
                    items:[{
                        id:'mainpanel',
                        autoWidth:true,
                        margins:'0 0 0 0',
                        layout:'anchor',
                        layoutConfig:{animate:true}
                        
                        },{
                        
                        html:'<br><br><br><br><img src=img\\biglogo.gif><br><br><h2>Welcome to Jello 5 Control Center</h2><br>'+txtVersion+': <b><i>'+jelloVersion+'</i></b><br><br>'+txtInWelcmsg, 
                        id:'mainpanel2',
                        //title: txtHomeStart1,
                        autoWidth:true,
                        style:'text-align:center;padding-top:5px;',
                        border:false,
                        margins:'5 5 5 5',
                        height:panelHeight,
                        layout:'anchor',
                        layoutConfig:{animate:true}
                        
                        }]
                        
                });
mainMenu();                
}

function mainMenu()
{
    tbar1 = new Ext.Toolbar({style:'padding-left:120px;margin-top:0;text-align:center;'});
    
        tbar1.add('-',{
        text:txtInsHTARun,
        id:'instrun',
        cls:'olrunicon',
        width:150,
        enableToggle:false,
			  listeners:{click:function(){pRun('jello5.hta');}}
        },'-',{
        text:'<b>'+txtInOLSetUp+'</b>',
        id:'instsetup',
        cls:'olseticon',
        width:150,
        enableToggle:true,
        toggleGroup:'itype',
			  handler:pInstall
        },'-',{
        text:'<b>'+txtInSetLoca+'</b>',
        id:'instloc',
        cls:'olsetloc',
        width:150,
        enableToggle:true,
        toggleGroup:'itype',
			  handler:pSettings
        },'-',
        {
        text:'<b>Import/Export</b>',
        id:'impexp',
        cls:'olsetexp',
        width:150,
        enableToggle:true,
        toggleGroup:'itype',
			  handler:pImpExp
        },'-',
        {
        text:txtMenuUtilsLinks,
        id:'instutls',
        cls:'olsetutl',
        width:150,
        enableToggle:true,
        toggleGroup:'itype',
			  handler:pUtils
        },'-'
        
);

tbar1.render('mainpanel');
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
        height:250,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        id:'tagform',
        buttonAlign:'center',
        width:650,
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
        width:650,
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
var furi=uri+fl;
furi='"'+furi+'"';
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


function pUtils()
{
var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");
ret="<br><br><img src=img//biglogo.gif><br><br><h1>"+txtMenuOtherFunctions+"</h1><br>";
ret+="<a class=jellolinktop onclick=pRun('jello5.hta')>Run Jello without Outlook (jello5.hta)</a><br>";
ret+="<a class=jellolinktop onclick=pRun2('jello5.htm')>Run Jello in Internet Explorer</a><br><br>";
ret+="<a class=jellolinktop onclick=pRun('help.pdf')><b>"+txtHelp+"</b> (help.pdf)</a><br><br>";
ret+="<hr><br><h1>Links</h1><br>";
ret+="<a class=jellolinktop href='http://jello-dashboard.net'><b>Jello Dashboard home page</b></a><br>";
ret+="<a class=jellolinktop href='http://jello-dashboard.net/forum'>Jello Forum</a><br>";
ret+="<a class=jellolinktop href='http://code.google.com/p/jello-dashboard/'>Jello Dashboard Open Source project at Google Code</a><br><br><hr>";
ret+="<br><br>2006-2010 Nicolas Sivridis (Licensed under Open Source Lisence GPLv3)";
mp.getEl().update(ret);
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
  furi=furi.replace("\/install.hta","");
  }
ret="<br><br><img src=img//biglogo.gif><br><div style=background:gainsboro;><br><h1>File Operations</h1><br><hr><br><h2>Export Settings to file</h2>";
ret+="<input class=htmlControls id=expfiled type=text size=100 value='"+furi+"'><br>"
ret+="<input class=htmlControls type=button onclick=exportFile() value='Export Settings'><br>";
ret+="<hr><br><h2>Import Settings from file</h2>";
ret+="<input type=file class=htmlControls id=impfiled value='"+furi+"' size=90 value='Import Settings'><br>";
ret+="<input type=button class=htmlControls onclick=importFile() value='Import Settings'><br><br>";
ret+="</div>";
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
