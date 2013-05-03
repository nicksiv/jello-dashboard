// Widgets

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


var wStore=new Array();
var twTitle=new Array();
var wJInboxActive=new Array();
var wJInboxActiveItem=new Array();
var wJInboxSecured=new Array();
var wJInboxClipbrd=new Array();


var mailStore=new Ext.data.SimpleStore({fields: [{name: 'entryid'},{name: 'sender'},{name: 'body'},{name: 'cc'},{name: 'bcc'}]});
var mailRecord = Ext.data.Record.create([{name: 'entryid'},{name: 'sender'},{name: 'body'},{name: 'cc'},{name: 'bcc'}]);

var widgetMap = [
        ['IMP',txtWidTImportant,'jImportant()',true,true,false,txtWdImp,'Yes|||Yes','icon_alert.gif'],
        ['WFL',txtWidTJMenu,'jelloMenu()',true,false,true,txtWdJWL,null,'j-icon.gif'],
        ['TIK',txtTicklers,'wTicklers()',false,false,false,txtWdTik,'D','ticklers16.png'],
        ['TAG',txtWidTJTagAct,'wTagActions()',false,false,false,txtWdTag,getActionTagName(),'list_components.gif'],
        ['PRJ',txtWidProjects,'wProject()',false,false,false,txtWdProj,'Yes|||No','project.gif'],
        ['OLW',txtOLView,'wOutlookView()',false,false,false,txtWdOLV,jello.inboxFolder,'folder.gif'],
        ['INB',txtwJDInbox,'wJDInbox()',false,false,false,txtWdOLV,'jello.inboxFolder|||200','inbox16.png'],
        ['SFO',txtwOLSearch,'wFolSearch()',false,false,false,txtwOLSearchDesc,jello.contactFolder,'sfolder.gif'],
        ['PIT',txtWidTJPostit,'wPostIt()',false,false,false,txtWdPostit,'PaleGoldenrod|||Verdana|||10','note.gif'],
        ['WEB',txtWidTJWPage,'wPage()',false,false,false,txtWdWebP,'Web page|||about:blank','page_url.gif'],
        ['HTM',txtWidTJHLFrag,'wInfo()',false,false,false,txtWdHTML,txtWidTJHLFrag+'|||<span class=caltime>'+txtWdHTMLdef+'...</span>','icon_extension.gif']
        ];

var widgetMapStore=new Ext.data.SimpleStore({
        fields: [
           {name: 'id'},
           {name: 'name'},
           {name: 'function'},
		   {name: 'unique',type:'boolean'},
		   {name: 'noremove',type:'boolean'},
		   {name: 'nocustomize',type:'boolean'},
		   {name: 'description'},
		   {name: 'defaults'},
		   {name: 'icon'}
        ],
        data:widgetMap
    });

function jelloMenu()
{
//menu widget renderer
var wret="new Ext.Panel({id:'wc00',frame:false,margins:'5 5 5 5',html:jelloMenuHTML()})";
return wret;
}

function jelloMenuHTML()
{
//jello menu widget
var ret="<table cellpadding=5 cellspacing=5 width=100%>";
ret+="<tr><td width=30 align=center><img src=img\\collect16.png></td><td><a onclick=pCollect(); class=jellolink><b>"+txtCollect+"</b></a></td><td><span class=widgettext>"+txtCollectDetails+"</span></td></tr>";
ret+="<tr><td align=center><img src=img\\inbox16.png></td><td><a onclick=pInbox(); class=jellolink><b>"+txtInbox+"</b></a> <span class=fkey>("+countInboxItems()+")</span></td><td><span class=widgettext>"+txtInboxDetails+"</span></td></tr>";
//ret+="<tr><td align=center><img src=img\\review.gif></td><td><a onclick=pReview(); class=jellolink><b>"+txtReview2+"</b></a></td><td><span class=widgettext>Review all actions by tag</span></td></tr>";
ret+="<tr><td align=center><img src=img\\ticklers16.png></td><td><a onclick=pTicklers(); class=jellolink><b>"+txtTicklers+"</b></a></td><td><span class=widgettext>"+txtTicklersDetails+"</span></td></tr>";
ret+="<tr><td align=center><img src=img\\master16.png></td><td><a onclick=dynamicLoad('masterlist.js','pMaster()'); class=jellolink><b>"+txtMaster+"</b></a></td><td><span class=widgettext>"+txtMasterDetails+"</span></td></tr>";
//ret+="<tr><td align=center></td><td colspan=2>&nbsp;</td></tr>";
ret+="<tr><td align=center><img src=img\\settings16.png></td><td><a onclick=dynamicLoad('settings.js','mngSettings(0)'); class=jellolink><b>"+txtSettings+"</b></a></td><td><span class=widgettext>"+txtSettingsDetails+"</span></td></tr>";
ret+="<tr><td>&nbsp;</td><td class=jellolink align=left><b>Create</b></td><td><input id='quickwe' onkeydown='WQuickEntry()' title='"+txtQuickEntryTitle+"' style='width:85%;border:1px gray solid;font-size:10px;' type=text size=30>&nbsp;<a style='cursor:hand;' onclick=WQuickEntryInfo();><img title='"+txtQuickEntryHelp+"' src=img//icon_info.gif></a></tr>";
ret+="</table>";

return ret;
}


function jImportant(id)
{
//important widget renderer

//try{
var ws=getWidgetSettings(id);

  if(jImportant.arguments.length>1)
  {//customize
  if(!notEmpty(ws[0])){ws[0]="Yes";ws[1]="Yes";}
  var war = [new Ext.form.ComboBox({fieldLabel: txtShowTodayActions,id: 'itod',store:new Array('Yes','No'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[0], mode:'local'}),new Ext.form.ComboBox({fieldLabel: txtShowDueActions,id: 'ipd',store:new Array('Yes','No'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[1], mode:'local'})];
  return war;
  }
//}catch(e){}
  

var j="";
j=jImportantHTML(id,ws);
var wret="new Ext.Panel({id:'wc0',margins:'5 5 5 5',html:jImportantHTML("+id+")[0]})";

 setTimeout(function(){
  if (j[1]>0)
  {var pnl=Ext.getCmp("w"+id);
  try{pnl.setTitle(txtImportantTitle+" ("+j[1]+")");}catch(e){}}
  },3);

return wret;

}

function jImportantHTML(wid,ws)
{
//important widget HTML code
var ht=getWidgetHeight(wid);
var ret="<div class=widgetInner style='overflow-y:auto;height:"+ht+"px;'>";
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var ImportantPastDueCount=0;
//if (webJDversion!=jelloVersion && notEmpty(webJDversion)){ret+="<span class=widgetalert>A new Jello Dashboard version is available. "+webJDversion+" </font><a href='http://www.jello-dashboard.net/download/' target='_blank'><b>Download</b></a></span><hr style='color:white;border-bottom:1px gray dotted;'>";}
try{if (ws[0]=="Yes"){}}catch(e){var ws=getWidgetSettings(wid);}
if(!notEmpty(ws[0])){ws[0]="Yes";ws[1]="Yes";}
var nas=getActionTagName();
var emph="";
//today tasks
if (ws[0]=="Yes")
{
ret+="<span class=widgettext style=color:navy;padding-left:0px;height:20px;><b>Actions Today</b>&nbsp;</span><br>";
var its=iF.Restrict("@SQL=http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2 AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040='Today'");
  ImportantPastDueCount=its.Count;
  if(its.Count>0)
  {
    for (var x=1;x<=its.Count;x++)
    {
    var tList=its(x).itemProperties.item(catProperty).Value;
    if (notEmpty(tList)){tList="["+tList+"]";}
    if (tList.search(nas)>-1){emph="<b>";}else{emph="";}
    ret+="<span class=widgetaction><a class=jellolink onclick=scAction('"+its(x).EntryID+"')>"+emph+its(x).Subject+"</b> - " + DisplayAppDate(new Date(its(x).DueDate))+ "</a>&nbsp;<span class=caltime>"+tList+"</span></span><br>";
    }
  }
  else
  {
  ret+="<span class=fkey>"+txtNoItems+"</span><br><br>";
  }
}

//past due
if (ws[1]=="Yes")
{
if (ws[0]=="Yes"){ret+="<hr style='color:white;border-bottom:1px gray dotted;'>";}
ret+="<span class=widgettext style=padding-left:0px;height:20px;><b>"+txtWdImpAPD+"</b>&nbsp;</span><br>";
var its=iF.Restrict("@SQL=http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2 AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040<'Today'");
ImportantPastDueCount=ImportantPastDueCount+its.Count;
  if(its.Count>0){
    for (var x=1;x<=its.Count;x++)
    {
    var tList=its(x).itemProperties.item(catProperty).Value;
    if (notEmpty(tList)){tList="["+tList+"]";}
    if (tList.search(nas)>-1){emph="<b>";}else{emph="";}
    ret+="<span class=widgetaction><a class=jellolink onclick=scAction('"+its(x).EntryID+"')>"+emph+its(x).Subject+"</b> - " + DisplayAppDate(new Date(its(x).DueDate))+ "</a>&nbsp;<span class=caltime>"+tList+"</span></span><br>";
    }
  }
  else
  {
  ret+="<span class=fkey>"+txtNoItems+"</span><br><br>";
  }
}
 
if (ws[0]=="No" && ws[1]=="No"){ret+=txtNoDataMsg;}
 
if (jello.migrated==false)
{
dynamicLoad("migrate.js","loadMigrationScr()");
var sm=settingsMigration(true);
if (notEmpty(sm))
  {//if user did not migrated from previous jello settings show these
  ret+="<br><span class=widgetalert><b>"+sm+"</b></span><br>";
  }
}

if (notEmpty(jello.archiveFolder)==false && jello.noUseMail==0)
{ret+="<br><span class=widgetalert>"+txtWdImpNoArc.replace("%1","<a class=jellolink onclick=mngSettings(0);>"+txtSettings+"</a>")+"</span><br>";}
var rv=new Array();
rv.push(ret);
rv.push (ImportantPastDueCount);
return rv;
}

function refreshImportant(id)
{
var ws=getWidgetSettings(id);
var j=jImportantHTML(id,ws);
Ext.getCmp("wc0").getEl().update(jImportantHTML(id,ws)[0]);
  if (j[1]>0)
  {var pnl=Ext.getCmp("w"+id);
  try{pnl.setTitle(txtImportantTitle+" ("+j[1]+")");}catch(e){}}

}
function wInfo(id)
{
//Info Widget
//Show a HTML code fragment
//Settings:0=Widget Title, 1=HTML code
var ws=getWidgetSettings(id);

  if(wInfo.arguments.length>1)
  {//customize
  var war = [{fieldLabel: 'Title',id: 'c0',value:ws[0],allowBlank:false},new Ext.form.TextArea({id: 'c1',height:180,emptyText:'Enter HTML code here',value:ws[1], fieldLabel:'HTML Code'})];
  return war;
  }

var ret="";
//set widget's body
ws[1]=ws[1].replace(/'/g, '"');
ret="new Ext.Panel({id:'wc"+id+"',margins:'2 2 2 2',html:'<div style=padding:10 10 10 10;>"+ws[1]+"</div>'})";
//set widget's title
setTimeout(function(){Ext.getCmp("w"+id).setTitle(ws[0]);},3);
return ret;
}

function wTicklers(id)
{
//Tickler Widget
//Show Ticklers for selected period
//Settings:0=Time period (D/W/M/A)
// 1 arg means create widget for home
// 2 args when we're doing settings
// 3 args for popups, if arg 3 =="", popup is for home page popup

if( wTicklers.arguments.length <3 || wTicklers.arguments[2] == "" )
	var ws=getWidgetSettings(id);
else{
	// popup and not widget
	id = 0;
}


  if(wTicklers.arguments.length==2)
  {//customize
  var dPeriods = [['D',txtDay],['W',txtWeek],['M',txtMonth],['A',txtAgenda]];
  var periodStore=new Ext.data.SimpleStore({fields: [{name: 'value'}, {name: 'text'}],data:dPeriods});
  var war = [new Ext.form.ComboBox({fieldLabel:txtWdTikTmPer ,id: 'c0',store:periodStore,editable:false,forceSelection:true,displayField:'text',valueField:'value',triggerAction:'all',mode:'local',value:ws[0]})];
  return war;
  }

//set widget's contents
// set store assuming widget and not popup
var storeID = "wStore["+id+"]";
var gid = "wgrid:"+id;
var dMenu ="";

if(wTicklers.arguments.length==1 || (wTicklers.arguments.length==3 &&
	wTicklers.arguments[2] == ""))
{
	if (wTicklers.arguments.length==3 && wTicklers.arguments[2] == "")
		wStore[id] = pTicklers(ws[0],true);
	else
		wStore[id] = pTicklers(ws[0]);
}else
{
// plain popup not widget or home page popup
if( ibtStore == null)
	ibtStore = pTicklers('D', true);
storeID = "ibtStore";
gid = "ibxtgrid:0";
dMenu="{icon:'img//calendar.gif',cls:'x-btn-icon',id:'popdate',tooltip:txtTikGoDate, menu:dateMenu, menuAlign:'c-tl?'},";
}
ret=[
	"new Ext.grid.GridPanel({store:"+storeID+",id:'"+gid+"',",
     "tbar:[{icon:'img//appoint.gif',cls:'x-btn-icon',id:'vnew',",
     "tooltip: '<b>"+txtCreate+"</b><br>"+txtTikCrNew+"',handler : ticklerSelected},",

	 "{icon: 'img//page_delete.gif',cls:'x-btn-icon',id:'vdel',tooltip: '"+txtDelItmInfo+"',",
     "handler : ticklerSelected},'-',{icon: 'img//check.gif',cls:'x-btn-icon',id:'vdone',tooltip: '"+txtCompleteItmInfo+"',",
     "handler : ticklerSelected},'->',",
     	 dMenu+"{text:'"+txtWdTikOpScr+"',listeners:{click:function(){pTicklers();}}}],columns: [{header: '', width: 5, fixed:true, sortable:false, renderer: getImportance, dataIndex: 'importance'},",
     "{header: '', width: 25, hideable:false, fixed:true, sortable: false, renderer: getIcon, dataIndex: 'icon'},",
     "{header: '"+txtTickler+"', width: 200, sortable: true, renderer: renderTickSubject,dataIndex: 'subject'},",
     "{header: '"+txtStartDate+"', width: 100, hidden:false, sortable: true, renderer:DisplayCalDate, dataIndex: 'due'},",
     "{header: '"+txtDuration+"', width: 50, hidden:true, sortable: true, renderer:DisplayDuration, dataIndex: 'duration'},",
     "{header: '"+txtLocation+"', width: 50, hidden:true, sortable: true, dataIndex: 'location'},",
     "{header: '"+txtPosition+"', width: 50, hidden:true, sortable: true, dataIndex: 'daypos'}",
     "],stripeRows: true,autoScroll:true,deferRowRender:false,enableColumnHide:true,viewConfig:{emptyText:'"+txtNoDispItms+"'},",
     "trackMouseOver:true,width:250,listeners:{mouseover: function(e){thisGrid='"+gid+"';},rowdblclick: function(g,row,e){openTicklerItem(null,g);}, cellcontextmenu: function (g, row, cell, e){rightClickItemMenu(e,row,g);}}})"].join("");


if(wTicklers.arguments.length!=2)
{
if (wTicklers.arguments[2]=="")
{
twTitle[id]=updateTicklerTitle(true);
  setTimeout(function(){
  var pnl=Ext.getCmp("w"+id);
  pnl.setTitle(txtTicklers+": "+twTitle[id]);

      try{
      var g=Ext.getCmp("wgrid:"+id);
      g.on('columnresize',function(index,size){saveGridState("wgrid:"+id);});
      g.on('sortchange',function(){saveGridState("wgrid:"+id);});
      g.getColumnModel().on('columnmoved',function(){saveGridState("wgrid:"+id);});
      g.getColumnModel().on('hiddenchange',function(){saveGridState("wgrid:"+id);});
      restoreGridState("wgrid:"+id);
      }catch(e){}
  },3);
}
}
return ret;
}

function wTagActions(id)
{//Tag Actions Widget
//Show Actions of selected tag
//Settings:0=Tag name
var ws=getWidgetSettings(id);

if(wTagActions.arguments.length==5)
{
//set only title
  setTimeout(function(){var pnl=Ext.getCmp("w"+id);pnl.setTitle(ws[0]);});
  return;
}

  if(wTagActions.arguments.length==2)
  {//customize
  var war = [new Ext.form.ComboBox({fieldLabel: txtTag,id: 'itags',store:globalTags,hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[0], emptyText:txtTagToList,mode:'local'})];
  return war;
  }

//set widget's contents

if(wTagActions.arguments.length==1)
{try{wStore[id] = showContext(ws[0],2);}catch(e){ret="new Ext.Panel({html:'&nbsp;ERROR!&nbsp;<b>"+ws[0]+"</b>"+txtTagNotFound+"'})";return ret;}}


ret=["new Ext.grid.GridPanel({store:wStore["+id+"],id:'wgrid:"+id+"',",
	"tbar:[{icon: 'img//new.gif',cls:'x-btn-icon',tooltip:'"+txtAddNewItmInfo+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();editAction(null,null,null,'"+ws[0]+"');}}},",
	"{icon: 'img//page_edit.gif',cls:'x-btn-icon',tooltip: '"+txtEdSelItmInf+"',id:'aedit',handler : actionSelected},",
	"{icon: 'img//page_delete.gif',cls:'x-btn-icon',id:'adel',tooltip: '"+txtDelItmInfo+"',handler : actionSelected},'-',",
	"{icon: 'img//IsNext.gif',cls:'x-btn-icon',id:'anext',tooltip: '"+txtSetNAInfo+"',handler : actionSelected},",
	"{icon: 'img//check.gif',cls:'x-btn-icon',id:'adone',tooltip: '"+txtCompleteItmInfo+"',handler : actionSelected},",
	"new Ext.form.ComboBox({id: 'tagcmbw"+id+"',store:globalTags,hideTrigger:false,typeAhead:false,triggerAction:'all',width:80,listWidth:220,emptyText:txtTagSelInfo,mode:'local',maxHeight:250,listeners:{specialkey: function(f, e){if(e.getKey()==e.ENTER){actionSelected('addtag','w"+id+"');userTriggered = true;e.stopEvent();f.el.blur();Ext.getCmp('tagcmbw"+id+"').reset();}},select: function(cb,rec,idx){actionSelected('addtag','w"+id+"');Ext.getCmp('tagcmbw"+id+"').reset();}}}),",
	"{tooltip:txtSetPrivacy,cls:'x-btn-icon',id:'alock',icon: 'img//page_lock.gif',handler:actionSelected},",
	"{tooltip:txtOpenInOutlook,cls:'x-btn-icon',id:'aoutlook',icon: 'img//info.gif',handler:actionSelected}],",
	"columns: [{header: '', width: 30, hideable:false, fixed:true, sortable: false, renderer: getIcon, dataIndex: 'icon'},",
    "{header: '"+txtCreated+"', width: 80, hidden:true, sortable: true, renderer:DisplayAppDate, dataIndex: 'created'},",
    "{header: '"+txtSubject+"', width: 200, sortable: true, renderer: renderSubject,dataIndex: 'subject'},",
    "{header: '"+txtDueDate+"', width: 80, sortable: true, renderer:DisplayAppDate, dataIndex: 'due'},",
    "{header: '"+txtTags+"', width: 80,hidden:true, sortable: true, dataIndex: 'tags'},",
	"{header: '"+txtContacts+"', width: 80,hidden:true, sortable: true, dataIndex: 'contacts'},",
    "{header: '"+txtNotes+"', width: 80,hidden:true, sortable: true, dataIndex: 'body'}],",
	"stripeRows: true,autoScroll:true,enableColumnHide:true,deferRowRender:true,viewConfig:{emptyText:'"+txtNoDispItms+"'},",
	"trackMouseOver:true,showPreview:true,width:250,listeners:{mouseover: function(e){thisGrid='wgrid:"+id+"'},rowdblclick: function(g,row,e){openGridRecord(g);}",
	"},enableColumnMove:false,border:false})"].join("");

if(wTagActions.arguments.length==1)
{
  setTimeout(function(){
  var pnl=Ext.getCmp("w"+id);
  pnl.setTitle(ws[0]);
  try{
  var g=Ext.getCmp("wgrid:"+id);
         	g.getSelectionModel().on('rowselect', function(sm, rowIdx, r) {
		var detailPanel = Ext.getCmp('previewpanel');
		actionTpl.overwrite(detailPanel.body, r.data);

	   });
  g.on('columnresize',function(index,size){saveGridState("wgrid:"+id);});
  g.on('sortchange',function(){saveGridState("wgrid:"+id);});
  g.getColumnModel().on('columnmoved',function(){saveGridState("wgrid:"+id);});
  g.getColumnModel().on('hiddenchange',function(){saveGridState("wgrid:"+id);});
  try{g.getSelectionModel().selectRow(0);g.getView().focusRow(0);g.focus();}catch(e){}
  restoreGridState("wgrid:"+id);
  }catch(e){}

  },3);
}
return ret;
}

function wPostIt(id)
{
//Post it Widget
//Show notes
//Settings:0=color, 1=font, 2=size;
var wid = "wgrid:"+id;
if(wPostIt.arguments.length==3)
	wid = wPostIt.arguments[2];
if( id != -1)
	var ws=getWidgetSettings(id);
else{
	var ws = new Array();
	ws[2] = 10;
	ws[1]="Courier";
	ws[0] = 'PaleGoldenrod';
}

if (ws[2]<10){ws[2]=10;}
if (!notEmpty(ws[1])){ws[1]="Courier";}
if (!notEmpty(ws[2])){ws[2]=10;}

  if(wPostIt.arguments.length==2)
  {//customize
  var war = [new Ext.form.ComboBox({fieldLabel: txtWdPoColor,id: 'icol',store:new Array('Yellow','PaleGoldenrod','Linen','LightCyan','MistyRose','LightGreen','Gainsboro','NavajoWhite'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[0], mode:'local'}),new Ext.form.ComboBox({fieldLabel: txtWdPoFont,id: 'ifont',store:new Array('Courier','Tahoma','Verdana','Trebuchet MS', 'Arial', 'Sans-serif'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[1], mode:'local'}),new Ext.form.NumberField({fieldLabel:txtWdPoFontSz,value:ws[2], width:200,invalidText:'The text entered is invalid',id:'isize'})];
  return war;
  }

//get saved text
var pitvl="";
if(id != -1){
var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);

	if (r!=null)
	{
		try{pitvl=r.get("content");}catch(e){pitvl="not found";}
	}
}

//set widget's contents
ret=["new Ext.form.TextArea({id: '"+wid+"',width:250, emptyText:'"+txtWdPoNoteHr+"', style:'background:"+ws[0]+";font-family:"+ws[1]+";font-size:"+ws[2]+";',",
	 "enableKeyEvents:true,listeners:{keyup: function(t, e){if(e.getKey()==e.ENTER){wPostItSave(t,'"+id+"');}}}})"].join("");


//if(wPostIt.arguments.length==1)
if(wPostIt.arguments.length!=2)
{
  setTimeout(function(){
  try{
Ext.getCmp(wid).setValue(pitvl);
}catch(e){}
  },3);
}

return ret;
}

function wPostItSave(t,id)
{

//postIt Widget save routine
// get the text
var pitvl=t.getValue();
// if from popup and not postit widget
// save to temp string
if( id == -1){
	postItPopContent = pitvl;
	t.getEl().fadeIn();
	return;
}

var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);
	if (r!=null)
	{
		r.beginEdit();
		r.set("content",pitvl);
		r.endEdit();
		syncStore(widgetStore,"jello.widgets");
		jese.saveCurrent();
	}
t.getEl().fadeIn();

}

function wOutlookView(id)
{
//Outlook view widget
//Settings:0=Outlook folder ID, 1=Outlook view

var ws=getWidgetSettings(id);
var folder=setAndCheckArcFolder(ws[0]);

  if(wOutlookView.arguments.length==2)
  {//customize

  var olvws=new Array();
  for (var x=1;x<=folder.Views.Count;x++)
  {olvws.push(folder.Views.Item(x));}

  var war = [new Ext.form.TriggerField({fieldLabel:txtOLFolder,value:folder.FolderPath,width:200,onTriggerClick:function(e){wOutlookViewSelectFolder();},id:'folfolder'}),new Ext.form.Hidden({value:folder.EntryID,id:'folfolderID'}),new Ext.form.ComboBox({fieldLabel: 'View',id: 'ffolderview',store:olvws,hideTrigger:false,typeAhead:false,editable:false,triggerAction:'all',value:ws[1], emptyText:txtWdOLSelVw,mode:'local'})];
  return war;
  }


if (folder!=null)
{
//set widget's contents
var comboMax=100;

//create tags list
tagStore.filter("istag",true);
widgetTags=[];
		for (var x=0;x<tagStore.getCount();x++)
	{
		var tn=tagStore.getAt(x).get("tag");
		var isArc=tagStore.getAt(x).get("archived");
    if (notEmpty(tn) && isArc!=true){widgetTags.push(Ext.util.Format.htmlEncode(tn));}
	}
widgetTags.sort();
tagStore.clearFilter();


var isActiveX=false;if (conStatus=="Outlook ActiveX"){isActiveX=true;}
var olpre="";
//do not display preview pane button in old outlook versions and in HTA/IE mode
if (!isActiveX && OLversion>10){olpre="{icon:'img//olpreview.gif',cls:'x-btn-icon',id:'oprevw',hidden:"+isActiveX+",tooltip:'"+txtOLPwPaneTgl+"',listeners:{click: function(b,e){toggleOLViewPreview();}}},";}

var oltbar=["tbar:["+olpre+"{icon: 'img//new.gif',cls:'x-btn-icon',tooltip: '"+txtNewHere+"',listeners:{click: function(b,e){addOLItemHere(null,'"+folder.EntryID+"');}}},",
   "new Ext.form.ComboBox({fieldLabel: '"+txtTags+"',width:90, listWidth:200, store:widgetTags,id: 'itags:"+id+"',hideTrigger:false,typeAhead:false,triggerAction:'all',emptyText:'"+txtTagSelInfo+"',mode:'local',",
   "maxHeight:"+comboMax+",listeners:",
   "{expand: function(c){lowerOLViewCtl(true,'olv:"+id+"');},collapse: function(c){lowerOLViewCtl(false,'olv:"+id+"');},select: function(cb,rec,idx){inboxAction(null,'tag','olv:"+id+"');}}}),",
   "{icon:'img//backup.gif',cls:'x-btn-icon',tooltip:'"+txtArcvInfo+"',width:40,id:'iarc',listeners:{click: function(b,e){inboxAction(b,null,'olv:"+id+"');}}},{icon: 'img//collect16.png',cls:'x-btn-icon',id:'arconly',tooltip:'"+ txtArchiveOnly+"',listeners:{click: function(b,e){inboxAction(b,null,'olv:"+id+"');}}},",
   "'-',{icon:'img//reply.gif',cls:'x-btn-icon',tooltip:'"+ txtReplyInfo +"',id:'ireply',handler:function(b,e){inboxAction(b,null,'olv:"+id+"')}},",
   "{icon: 'img//page_delete.gif',cls:'x-btn-icon',tooltip:'"+ txtDelete+"',handler:function(b,e){inboxAction(b,'del','olv:"+id+"')}},",
   "'->',{cls:'x-btn-icon',tooltip:'"+txtCrAppForInfo+"',icon: 'img//appoint.gif',handler:function(b,e){inboxAction(null,'cal','olv:"+id+"');}},",
   "{cls:'x-btn-icon',icon: 'img//user.gif',tooltip: '"+txtDelgForInfo+"',id:'ideleg',handler:function(b,e){inboxAction(b,null,'olv:"+id+"');}},",
   "{cls:'x-btn-icon',icon: 'img//list_review.gif',tooltip: '"+txtRevForInfo+"',id:'irev',handler:function(b,e){inboxAction(b,null,'olv:"+id+"');}},",
   "{cls:'x-btn-icon',icon: 'img//list_someday.gif',tooltip: '"+txtIncubForInfo+"',id:'iinc',handler:function(b,e){inboxAction(b,null,'olv:"+id+"');}},",
   "'-'],"].join("");

if (Browser.Version()>=7 && conStatus=="Outlook ActiveX"){oltbar="tbar:["+olpre+"{text: 'Browser does not support this',handler:function(b,e){alert('IE 7 now blocks handling of Outlook view control contents. You can manage contents only by the built in keyboard shortcuts or context menus.');}}],";}

ret=["new Ext.Panel({frame:false,border:false,id:'olview:"+id+"',",
   oltbar+"html:olviewCtl('"+folder.EntryID+"','olv:"+id+"'),",
   "autoWidth:true})"].join("");


}
else
{
//folder not exists! display an unavailable outlook folder message

}
if(wOutlookView.arguments.length==1)
{
//set widget title
  setTimeout(function(){
  var pnl=Ext.getCmp("w"+id);
  try{ttl=folder.FolderPath+" ("+folder.Items.Count+")";}catch(e){ttl="Undefined Folder";}
  if (ttl.length>25){ttl=ttl.substr(0,10)+"..."+ttl.slice((ttl.length-21),ttl.length);}
  try{pnl.setTitle(ttl);}catch(e){}
  try{var ov=document.getElementById("olv:"+id);ov.View=ws[1];}catch(e){}
  toggleOLViewPreview(false);
  },3);
}

return ret;
}

function wOutlookViewSelectFolder(noViews,allTypes)
{
//tag folder selection dialog
var t=NSpace.PickFolder();
if (noViews && t.DefaultItemType!=0 && !allTypes){alert(txtMailFolderPrompt);return;}
	if (t!=null )
	{
  //must be a mail folder
  Ext.getCmp("folfolderID").setValue(t.EntryID);
  Ext.getCmp("folfolder").setValue(t.FolderPath);

 if (noViews==null || noViews==false)
 {
  var olvws=new Array();
  for (var x=1;x<=t.Views.Count;x++)
  {olvws.push(t.Views.Item(x));}
  Ext.getCmp("ffolderview").store.loadData(olvws);
 }
}

}

function wPage(id)
{
//Web page Widget
//Show a web page in a widget
//Settings:0=Widget Title, 1=URL
var ws=getWidgetSettings(id);

  if(wPage.arguments.length>1)
  {//customize
  var war = [{fieldLabel: txtWdHTMLTtl,id: 'c0',value:ws[0],allowBlank:false},{fieldLabel: txtWdHTMLUrl+' http://...',id: 'c1',value:ws[1],allowBlank:false}];
  return war;
  }

var ret="";
//set widget's body
ret="new Ext.Panel({id:'wc"+id+"',margins:'2 2 2 2',html:'<iframe src="+ws[1]+" width=100% height=100% style=border:none;padding-top:3px;padding-left:3px;>'})";
//set widget's title
setTimeout(function(){Ext.getCmp("w"+id).setTitle(ws[0]);},3);
return ret;
}

var ibtStore = null;// = new Array();
var haveTicklerPopupCache = false;


function ticklerPopup()
{

	//ibtStore = new Array();
	var fn = wTicklers(null, null, "M");
	var dateMenu = new Ext.menu.DateMenu({
        handler : function(dp, date){
        var a=new Date(date);var b=new Date(date);b.setDate(b.getDate() + parseInt(jello.ticklerPopupDuration));
		getTicklerFile(a,b,1,Ext.getCmp('ibxtgrid:0').getStore(),true,true);
        }
    });
	var ticklerDisp=eval(fn);
	// wTicklers is set to return a day, now do 2 weeks
	 if( haveTicklerPopupCache == false){
		 var a=new Date();var b=new Date();b.setDate(b.getDate() + parseInt(jello.ticklerPopupDuration));
		getTicklerFile(a,b,1,Ext.getCmp('ibxtgrid:0').getStore(),true,true);
		haveTicklerPopupCache = true;
	}

    var win = new Ext.Window({
        title: txtTicklers,
        width: 480,
        height:370,
        id:'inboxticklerwidgetform',
        minWidth: 300,
        minHeight: 280,
        resizable:false,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: ticklerDisp,
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();}
            }
        }]
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);

}

var postItPopContent="";

function postItPopup()
{
	var sid = -1;
	for(var i=0; i < widgetStore.getCount(); i++)
	{
		var rec = widgetStore.getAt(i);
		if( rec.get("id") == "PIT"){
			sid = rec.get("widgetID");
			break;
		}

	}
	var wid = wPostIt(sid,null,"posttext");
	var elem = eval(wid);
	var pCol={
        id:'save',
        text:'Collect',
        tooltip:'Collect',
        handler: function(e, target, panel){
        ;
		var coltxt=Ext.getCmp("posttext").getRawValue();
            var c=importText(true,coltxt);
            if (c==true){importText(false,coltxt);win.destroy();}
			win.destroy();
        }
    };

	var win = new Ext.Window({
        title: txtSimpleCollect,
        width: 480,
        height:370,
        id:'postitpop',
        minWidth: 300,
        minHeight: 280,
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
		tools: [pCol],
        items: elem,
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text: txtBtnColText,
            tooltip:txtCollectBtnInfo,
            iconCls:'colectbuticon',
            listeners:{
            click: function(b,e){
       		var coltxt=Ext.getCmp("posttext").getRawValue();
            var c=importText(true,coltxt);
            if (c==true){importText(false,coltxt);win.destroy();return;}
            win.destroy();
            }
            }

        },{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();}
            }
        }]
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);

}


function homePopup()
{
	var isup = Ext.getCmp("homepop");
	if( typeof(isup) != "undefined" && isup != null){
		showOLViewCtl(false);
		isup.show();
		isup.doLayout();
		isup.setActive();
		return;
	}
	var x = pHome(true);
	if( x == null)
		return;

var win = new Ext.Window({
        title: txtHome,
        width: 600,
        height: 500,
        id:'homepop',
        minWidth: 300,
        minHeight: 280,
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        defaultButton: 0,
        minimizable: true,
        items: x,
        modal:true,
        listeners: {
		destroy:function(t){showOLViewCtl(true);},
		minimize: function(t){win.hide();}},
        buttons: [{
            text: txtSBHide,
            tooltip:txtSBHide,
            listeners:{
            click: function(b,e){
            win.hide();}
            }
        	}
            ]
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);
}

function wProject(id)
{
//Projects Widget
//Show active and inactive projects and their progress
//Settings:0=Widget Title, 1=Show inactive projects
var ws=getWidgetSettings(id);

  if(wProject.arguments.length>1)
  {//customize
  var war = [new Ext.form.ComboBox({fieldLabel: txtWdProjInac,id: 'icol',store:new Array('No','Both Active and inactive','Only Inactive'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[0], mode:'local'}),new Ext.form.ComboBox({fieldLabel: txtWdProjPath,id: 'ipth',store:new Array('Yes','No'),hideTrigger:false,typeAhead:false,triggerAction:'all',value:ws[1], mode:'local'})];
  return war;
  }

var ret="";
//set widget's body
ret="new Ext.Panel({id:'wc"+id+"',autoScroll:true,margins:'2 2 2 2',bodyStyle:'padding:0px 5px 5px 5px',html:wProjectsContent('"+ws[0]+"','"+ws[1]+"',"+id+")})";
//set widget's title
setTimeout(function(){try{Ext.getCmp("w"+id).setTitle(txtWidProjects);}catch(e){}},3);
return ret;
}

function wProjectsContent(showInactive,showPath,id)
{
var filteron=txtActive;
var showInac=false;if (trimLow(showInactive).substr(0,4)=="only"){showInac=true;filteron=txtInactive;}
var showBoth=false;if (trimLow(showInactive).substr(0,4)=="both"){showBoth=true;filteron=txtAll;}
var showPth=false;if (trimLow(showPath)=="yes"){showPth=true;}

pwStore=tagStore;  var tgCount=0;
var ret="<table cellspacing=0 cellpadding=0>";

 var pathTag=function(r){
    var tname=r.get("tag");
    var par=r.get("parent");
    var caption=Trim(TagFullPath(par,tname,false));
    r.beginEdit();r.set("path",caption);r.endEdit();
    };

    var renderTag = function(r){
      var tname=r.get("tag");
      var tid=r.get("id");
      var cls=r.get("closed");
      var ispj=r.get("isproject");
      if (!ispj){return;}
      if (cls){return;}
      var tarc=r.get("archived");
      //if (tpv || tarc){return;}

      if (!showInac && !showBoth && tarc){return;}
      if (showInac && !tarc){return;}

      var tgname=tname.replace(new RegExp(" ","g"),"~");
      var caption="";
      if (showPth){caption=r.get("path");}else {caption=tname;}
      var tagLink="<a class=jellolinkreview onclick=showContext('"+tgname+"',1);>"+caption+"</a>";
      var tCount=showContext(tname,0,null,0);
      var tCounter=0;if (tCount>0){tCounter="("+tCount+")";}else{tCounter="";}
      var aall=0;aall=showContext(tname,0,null,4);
      var perc=100-((tCount*100)/aall);
      if (isNaN(perc)){perc="0";}
      perc=getPercentIcons(perc,tCounter,aall);
      ret+="<tr><td width=5% height=22 style='border-bottom:1px dotted gray;' valign='middle'><a style='cursor:hand;' onclick='editTag("+tid+")' title='"+txtEditTag+"'>"+actionListTitleIcon(tname)+"</a></td>";
      ret+="<td style='border-bottom:1px dotted gray;' valign=middle width=85%><span>"+tagLink+" <span class=fkey>"+tCounter+"</span></span>";
      ret+=pwTools(tid,tname)+"</span></td>";
      ret+="<td style='border-bottom:1px dotted gray;' width=5%><span>"+perc+"</span></td></tr>";
      tgCount++;
    };

if (showPth){pwStore.each(pathTag);}
pwStore.sort("path","ASC");
pwStore.each(renderTag);
//for (var x=tgCount;x<=3;x++){ret+="<td colspan=2>&nbsp;</td>";}ret+="</tr>";
ret+="</table>";
setTimeout(function(){try{Ext.getCmp("w"+id).setTitle(txtWidProjects+ "-"+filteron+" ("+tgCount+")");}catch(e){}},6);
return ret;

}

function pwTools(tid,tname)
{
//project icons for the project widget
var ret="";
//find NAs
var itms=showContext(tid,2,null,5);
var itco=itms.getCount();
	if (itco>0)
	{
	var nn=itms.getAt(0).get("subject");
	ret="&nbsp;<img style='cursor:hand;' title='"+nn+" ("+itco+" "+txtNextActions+")' src="+imgPath+icIsNext+">";
	}

//find waiting
var itms=showContext(tid,2,null,9);
var itco=itms.getCount();
	if (itco>0)
	{
	var nn=itms.getAt(0).get("subject");
	ret+="&nbsp;<img style='cursor:hand;' title='"+nn+" ("+itco+" "+txtWaiting+")' src='img\\list_waiting.gif'>";
	}

	return ret;
}

function WQuickEntry()
{
//quick entry
	if(event.keyCode==13)
	{
	var qe=document.getElementById("quickwe");
	var val=qe.value;
	var eType=0;
		if(val.substr(0,1)=="#"){eType=2;val=val.substr(1,500);}//journal
		if(val.substr(0,1)=="$"){eType=1;val=val.substr(1,500);}//tag
		if(val.substr(0,1)=="%"){eType=3;val=val.substr(1,500);}//project
		vv=val.split(",");
		val=Trim(vv[0]);
		var tg=null;
		if (vv.length>1){tg=trimLow(vv[1]);}


		var foundTags=new Array();
		//select tag
		if (tg!=null)
		{

			  var tgfind = function(r){
				var vl=trimLow(r.get("tag"));
				vl=Ext.escapeRe(vl);
				if (trimLow(vl).search(tg)>-1){foundTags.push(r.get("id"));}
			  };

			tagStore.each(tgfind);

			//one tag found
			if (foundTags.length==1){WQuickListAdd(val,eType,foundTags[0]);}
			else
			{
			//multiple tags
			WQuickSelector(foundTags,eType,val);
			}

		}
		else
		{
		//no tags
		WQuickListAdd(val,eType,0);
		}



	}
}

function WQuickSelector(foundTags,eType,val)
{
//quick entry tag selector window

	if (foundTags.length==0)
	{
		Ext.info.msg(txtQuickCcTitle, txtQuickCcMsg);
	    var del=document.getElementById("quickwe");
		del.focus;
		return;
	}

showOLViewCtl(false);
var ft="";
	for (var x=0;x<foundTags.length-1;x++)
	{
	var r=getTagByID(foundTags[x]);
	ft+="<option value='"+foundTags[x]+"'>"+Ext.util.Format.htmlEncode(r.get("tag"))+"</option>";
	}

	var ftt="<select style='border:1px solid black;font-family:verdana;font-size:10px;width:100%;height:350;' id=SELECTTAG size=15 onkeydown=WQuickListSelected("+eType+",'"+val.replace(new RegExp(" ","g"),"~")+"')>"+ft;
var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: 'Select Tag',
        height:370,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        labelWidth:480,
        id:'tagselectform',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',
       	html:ftt

    });

    var win = new Ext.Window({
        title: 'New Entry',
        width: 300,
        height:370,
        id:'thetagselectform',
        minWidth: 300,
        minHeight: 200,
        resizable:false,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:3px;',
        buttonAlign:'center',
        items: simple,
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text: txtClose,
            tooltip:txtClose+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();}
            }
        }]
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);
    setTimeout(function(){var ste=document.getElementById("SELECTTAG");ste.focus();},16);

    var map = new Ext.KeyMap('thetagselectform', {
    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
            try{win.destroy();}catch(e){}

            },
    scope: this});

}

function WQuickListSelected(etype,val)
{
 //selected tag from tag selectior window for quick entry
  	if(event.keyCode==13)
	{
	sval=document.getElementById("SELECTTAG").value;
	WQuickListAdd(val,etype,sval);
	}
}

function WQuickListAdd(val,etype,sval)
{
//add quick item to outlook
val=val.replace(new RegExp("~","g")," ");

	var r=getTagByID(sval);
		if (r!=null)
		 {
		 var tg=r.get("tag");
		 }else{tg="["+txtInbox+"]";}

	if (etype==0)
	{//action
    var newA=createActionOL(val);
	if (tg!="["+txtInbox+"]"){newA.itemProperties.item(catProperty).Value=tg;}
	 newA.Save();
	 Ext.info.msg(txtActionCreated, "Action <a class=jellolinktop style=text-decoration:underline; onclick=scAction('"+newA.EntryID+"')>"+val+"</a> "+txtAddedTo+" "+tg);
	}

	if (etype==2)
	{//journal
	var jtg=tg;
	if (jtg=="["+txtInbox+"]"){jtg=" ";tg=txtNoTag;}
	var je=addJournalInTag(val,jtg);
	Ext.info.msg(txtReferenceCreated, txtJournalEntry+" <a class=jellolinktop style=text-decoration:underline; onclick=olItem('"+je.EntryID+"')>"+val+"</a> added to "+tg);
	}

	if (etype==1)
	{//tag
	var jtg=tg;
	if (jtg=="["+txtInbox+"]"){jtg=" ";tg=txtNoParent;}
	if(!createTag(val,false,sval,false,true,false)){return;}
	var cval=val.replace(new RegExp(" ","g"),"~");
	Ext.info.msg(txtTagCreated, txtTag+" <a class=jellolinktop style=text-decoration:underline; onclick=showContext('"+cval+"',1)>"+val+"</a> added to "+tg);
	}

	if (etype==3)
	{//project
	var jtg=tg;
	if (jtg=="["+txtInbox+"]"){jtg=" ";tg=txtNoParent;}
	if(!createTag(val,false,sval,false,true,true)){return;}
	var cval=val.replace(new RegExp(" ","g"),"~");
	Ext.info.msg(txtProjectCreated, txtProject+" <a class=jellolinktop style=text-decoration:underline; onclick=showContext('"+cval+"',1)>"+val+"</a> added to "+tg);
	}

 var del=document.getElementById("quickwe");
 del.value="";del.focus;
 try{Ext.getCmp("thetagselectform").destroy();}catch(e){}
}


function WQuickEntryInfo()
{
//information for the quick entry of the menu widget
//alert(txtQWdEntryIf);
           Ext.Msg.show({
           title:txtQWdEntryIfTitle,
           msg: txtQWdEntryIf,
           buttons: {ok:'OK'},
           fn: function(b,t){},
           animEl: 'elId',
           icon: Ext.MessageBox.INFORMATION
        });

}

function wJDInbox(id)
{
//Jello inbox widget
//Settings:0=Outlook inbox folder ID

var ws=getWidgetSettings(id);
var folder=setAndCheckArcFolder(ws[0]);
  if(wJDInbox.arguments.length==2)
  {//customize
  var ffpath="<No Folder>",ffid=0;
  try{ffpath=folder.FolderPath;ffid=folder.EntryID;}catch(e){}
  
  var war = [new Ext.form.TriggerField({fieldLabel:txtOLFolder,value:ffpath,width:200,onTriggerClick:function(e){wOutlookViewSelectFolder(true);},id:'folfolder'}),new Ext.form.Hidden({value:ffid,id:'folfolderID'}),new Ext.form.Field({fieldLabel:'Message Length',value:ws[1],id:'chlen'})];
  return war;
  }


if (folder!=null)
{
//set widget's contents
var comboMax=100;

if(typeof(wJInboxActive[id])=="undefined"){wJInboxActive[id]=1;}
//create tags list
tagStore.filter("istag",true);
widgetTags=[];
		for (var x=0;x<tagStore.getCount();x++)
	{
		var tn=tagStore.getAt(x).get("tag");
		var isArc=tagStore.getAt(x).get("archived");
    if (notEmpty(tn) && isArc!=true){widgetTags.push(Ext.util.Format.htmlEncode(tn));}
	}
widgetTags.sort();
tagStore.clearFilter();

/*"new Ext.SplitButton({icon:'img\\reply.gif',cls:'x-btn-icon',tooltip: '"+txtReply+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(1,"+id+");}}},",
"menu:new Ext.menu.Menu({items: [",
"{icon: 'img//forward.gif',cls:'x-btn-icon',tooltip: '"+txtFwd+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(3,"+id+");}}},",
"{icon: 'img//copy.gif',cls:'x-btn-icon',tooltip: '"+txtReplyAll+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(2,"+id+");}}}],",
"]})}),",
*/
    

var ret=["new Ext.Panel({id:'wci00',frame:false,margins:'5 5 5 5',",
"tbar:[{icon: 'img//action_back.gif',cls:'x-btn-icon',tooltip:'"+txtPrevious+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxGo(-1,"+id+");}}},",
"{icon: 'img//action_forward.gif',cls:'x-btn-icon',tooltip: '"+txtNext+"',id:'jmnext',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxGo(1,"+id+");}}},",
"'-',{icon: 'img//info.gif',cls:'x-btn-icon',tooltip: '"+txtOpenInOutlook+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(0,"+id+");}}},",
"{icon: 'img//page_delete.gif',cls:'x-btn-icon',tooltip: '"+txtDelete+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(4,"+id+");}}},",
"'-',new Ext.form.ComboBox({fieldLabel: '"+txtTags+"',width:90, listWidth:200, store:widgetTags,id: 'itags"+id+"',hideTrigger:false,typeAhead:false,triggerAction:'all',emptyText:'"+txtTagSelInfo+"',mode:'local',maxHeight:"+comboMax+",",
"listeners:{specialkey: function(f, e){if(e.getKey()==e.ENTER){userTriggered = true;e.stopEvent();f.el.blur();WJInboxFunction(6,"+id+");Ext.getCmp('itags"+id+"').reset();}},select:function(cb,rec,idx){WJInboxFunction(6,"+id+");}}}),",
"{iconCls:'olarchive',cls:'x-btn-icon',tooltip: '"+txtArchive+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(7,"+id+");}}},",
"'-',{icon: 'img//collect16.png',cls:'x-btn-icon',tooltip: '"+txtArchiveOnly+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(5,"+id+");}}},",
"{icon:'img//move.gif',cls:'x-btn-icon',tooltip: '"+txtMoveInfo+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(8,"+id+");}}},",
"{cls:'x-btn-icon',icon:'img//calendar.gif',tooltip:'"+txtDueForInfo+"',menu:",
"new Ext.menu.DateMenu({handler : function(dp, idate){var dt=idate.toUTCString();WJInboxFunction(12,"+id+",dt);}})},",
"'-',new Ext.SplitButton({icon: 'img//reply.gif',cls:'x-btn-icon',tooltip: '"+txtReply+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(1,"+id+");}},",
"menu:new Ext.menu.Menu({items: [",
"{icon: 'img//reply.gif',cls:'x-btn-icon',text:'"+txtReply+"',tooltip: '"+txtReply+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(1,"+id+");}}},",
"{icon: 'img//forward.gif',cls:'x-btn-icon',text:'"+txtFwd+"', tooltip: '"+txtFwd+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(3,"+id+");}}},",
"{icon: 'img//copy.gif',cls:'x-btn-icon',text:'"+txtReplyAll+"',tooltip: '"+txtReplyAll+"',listeners:{click: function(b,e){e.stopEvent();e.preventDefault();e.stopPropagation();WJInboxFunction(2,"+id+");}}}",
"]})})],html:renderJDInbox("+id+",'"+ws[0]+"')})"].join("");

}
else
{
var ret="new Ext.Panel({id:'wci00',frame:false,margins:'5 5 5 5',bodyStyle:'padding:0px 5px 5px 5px',html:'<b>Inbox folder does not exist</b>'})";
}

setTimeout(function(){
try{Ext.getCmp("w"+id).setTitle(txtInbox+" "+folder.FolderPath);
}catch(e){}},30);

return ret;
}

function renderJDInbox(iid,fid)
{
//render for the next inbox item
var ws=getWidgetSettings(iid);
var msgLength=200;
try{msgLength=ws[1];}catch(e){}
var ht=getWidgetHeight(iid);
var ret="<div id='jinbox"+iid+"'></div><div id='jinc"+iid+"' style='padding-left:5px;overflow-y:scroll;height:"+(ht-40)+";border-bottom:1px dotted black;'><table width=98% cellspacing=2>";
var num=wJInboxActive[iid];

//mail items
var jib=NSpace.GetFolderFromID(fid).Items;
var counter=jib.Count;

if (counter==0){ret="<p align=center><br><b>Inbox Empty!</b><br>wow...</p>";return ret;}

if (num==0){num=counter;wJInboxActive[iid]=num;}
if (num>counter){num=1;wJInboxActive[iid]=num;}

jib.Sort("[ReceivedTime]",true);

var it=jib(num);
//wJInboxActive[iid]++;
var dd=new Date(it.ReceivedTime);
var mailic="widgetmail";
if (it.Importance==2){mailic="widgetimpmail title='High Importance'";}
var atts=it.Attachments.Count;
var attLink="";
var unr=it.UnRead;
var flg=it.FlagStatus;
var msgCls="";if (unr){msgCls=" style='font-weight:bold'";}
var dp=getJDueProperty(it);
var itDue="";
if (notEmpty(dp)){itDue="<span onclick=WJInboxFunction(13,'"+iid+"'); title='"+txtDueDate+". Click to remove' class=tagnext style='cursor:hand;'>"+DisplayAppDate(new Date(dp))+"</span>";}
if (atts>0){attLink="&nbsp;<a style:'cursor:hand;' title='"+atts+" "+txtAttachments+"'><span style='font: x-small wingdings;font-size:17px;color:gray'>2</span>"+atts+"</a>";}
if (flg==2){attLink="&nbsp;<span style='font: x-small wingdings;font-size:17px;color:red'>P</span>"+attLink;}
ret+="<tr><td style='border-bottom:1px dotted black;'><span class=widgettext><b>"+num+"/"+counter+"</b></span></td>";
ret+="<td style='border-bottom:1px dotted black;' align=left width:90%><span class=widgettext>"+DisplayDate(dd)+"</span></td></tr>";
ret+="<tr><td colspan=2 align=left><span class="+mailic+msgCls+">"+Ext.util.Format.htmlEncode(it.Subject)+" "+attLink+" "+itDue+"</span></td></tr>";
var sender="?";var body="";var cc="";var bcc="";
var eid=it.EntryID;

    if (mailStore.getCount()>0)
    {
    mailStore.filter("entryid",new RegExp("^"+eid+"$"));
    var r=mailStore.getAt(0);
	sender=r.get("sender");
	body=r.get("body");
	cc="CC:"+r.get("cc");
	bcc="BCC:"+r.get("bcc");
	mailStore.clearFilter();
    }

    if (sender=="?")
    {
    if (jello.autoUpdateSecuredFields==false  || jello.autoUpdateSecuredFields==0 || jello.autoUpdateSecuredFields=="0")
    {
	body="<br><secured><a class=jellolinkTop style='font-size:10px;' onclick=WJInboxGetSecured("+iid+",'"+fid+"');>"+txtTryToGetBody+"</a></secured>";
    sender="?";
    }else
    {
    sender = it.SenderName;
	body=it.Body.substr(0,msgLength);
	cc="CC:"+it.CC;
	bcc="BCC:"+it.BCC;
    }
	}

ret+="<tr><td colspan=2><span class=widgetppl>"+sender+"</span></td></tr>";
ret+="<tr><td style='border-bottom:1px dotted black;'><span class=widgettext>To:</td>";
ret+="<td style='border-bottom:1px dotted black;'><span class=widgettext>"+cc+"<br>"+bcc+"</span></td></tr>";
ret+="<td colspan=2><span style='width:100%;overflow-y:none;' class=widgettext><i>"+body+"</i></span></td></tr>";
var tList=tagList(null,"[INBOX]",it.itemProperties.item(catProperty).Value,eid);
tList[0]=tList[0].replace(new RegExp("onclick=removeTag","g"),"onclick=removeTagJI");
tList[0]=tList[0].replace(new RegExp(eid+"'","g"),eid+"',"+iid);
ret+="</table></div>&nbsp;"+tList[0];
if (notEmpty(tList[0])){ret+="&nbsp;<a class=jellolinktop style='font-size:9px;' onclick=WJInboxFunction(10,"+iid+"); title='Click to copy tags from this message'>[CP]</a>";}
if (notEmpty(wJInboxClipbrd[iid])){ret+="&nbsp;<a class=jellolinktop style='font-size:9px;' onclick=WJInboxFunction(11,"+iid+"); title='Click to paste ["+wJInboxClipbrd[iid]+"] to this message'>[PS]</a>";}
ret+="<p align=center></p>";
wJInboxActiveItem[iid]=eid;
return ret;
}

function WJInboxGetSecured(id,fid)
{
//one time get secured data for jello inbox widget
var jib=NSpace.GetFolderFromID(fid).Items;
jib.Sort("[ReceivedTime]",true);
var jco=jib.Count;
	for (var x=1;x<=jco;x++)
	{
	var it=jib(x);
	var sndr=it.SenderName;
	var eid=it.EntryID;
	var mr=new mailRecord({entryid:eid,sender:sndr,body:it.Body.substr(0,200),cc:it.CC,bcc:it.BCC});
	mailStore.add(mr);
	status=x+"/"+jco+" "+it.Subject;
	}
	status="";
	doWidgetRefresh(Ext.getCmp("w"+id),id);
}

function WJInboxGo(nm,iid)
{
//navigate to jello inbox items
 wJInboxActive[iid];
	if (nm>0)
	{wJInboxActive[iid]++;}
	else
	{wJInboxActive[iid]--;}

	doWidgetRefresh(Ext.getCmp("w"+iid),iid);
}

function WJInboxFunction(fun,id,dt)
{
//jello inbox toolbar actions
var eid=wJInboxActiveItem[id];
	if (fun==0)
	{//open in outlook
	olItem(eid);
	}

	if (fun==1)
	{//reply
	var it=NSpace.GetItemFromID(eid);
	var nit=it.Reply;nit.Display();
	}

	if (fun==2)
	{//reply to all
	var it=NSpace.GetItemFromID(eid);
	var nit=it.ReplyAll;nit.Display();
	}

	if (fun==3)
	{//forward
	var it=NSpace.GetItemFromID(eid);
	var nit=it.Forward;nit.Display();
	}

	if (fun==4)
	{//delete
//	if (confirm(txtDelete+" ?"))
//		{
		var it=NSpace.GetItemFromID(eid);
		var nit=it.Delete();
		doWidgetRefresh(Ext.getCmp("w"+id),id);
//		}
	}

	if (fun==5)
	{//archive only
	var it=NSpace.GetItemFromID(eid);
	var its=new Array();
	its.push({olitem:it,record:null,update:false});
	doArchive(its,null,true);
	doWidgetRefresh(Ext.getCmp("w"+id),id);
	Ext.info.msg(txtwJDInbox,"Message <span class=caltime>"+it+"</span> archived");
	}

	if (fun==6)
	{//tag item
	var it=NSpace.GetItemFromID(eid);
	var tg=Ext.getCmp("itags"+id).getValue();
	    if (notEmpty(tg)==false)
		{var tg=Ext.getCmp("itags"+id).getRawValue();}
    if (notEmpty(tg)){if (addTagTo(tg,it,null,true)==false){return;}}
	doWidgetRefresh(Ext.getCmp("w"+id),id);
	}

	if (fun==7)
	{//archive item
	var it=NSpace.GetItemFromID(eid);
	var its=new Array();
	its.push({olitem:it,record:null,update:false});
	doArchive(its);
	doWidgetRefresh(Ext.getCmp("w"+id),id);
	Ext.info.msg(txtwJDInbox,"Message <span class=caltime>"+it+"</span> archived. Task created.");
	}

	if (fun==8)
	{//move item
	var it=NSpace.GetItemFromID(eid);
	var t=NSpace.PickFolder();
		if (t!=null)
		{
			if (t.DefaultItemType==0){

				try{moveMailTo(it,t,null,true);}catch(e){alert(txtMsgNoMoveErr);return;}
			}else
				try{copyMailTo(it,t);}catch(e){alert(txtMsgNoMoveErr);return;}
			imsg=txtMsgMoveTo+" <a class=jellolink style=text-decoration:underline; onclick=outlookView('"+t.EntryID+"')><b>"+t+"</b></a>";

		}


		doWidgetRefresh(Ext.getCmp("w"+id),id);
	}

	if (fun==10)
	{//copy tags
	var it=NSpace.GetItemFromID(eid);
	wJInboxClipbrd[id]=it.itemProperties.item(catProperty).Value;
	alert("Tags copied to clipboard.\n Click paste onto a message to use the same tags.");
	}

	if (fun==11)
	{//paste tags
	var it=NSpace.GetItemFromID(eid);
	it.itemProperties.item(catProperty).Value=wJInboxClipbrd[id];
	it.Save();
	doWidgetRefresh(Ext.getCmp("w"+id),id);
	}

 	if (fun==12)
	{//set due date
	var it=NSpace.GetItemFromID(eid);
  setDueDate(it,dt);
  doWidgetRefresh(Ext.getCmp("w"+id),id);
	}

 	if (fun==13)
	{//unset due date
	var it=NSpace.GetItemFromID(eid);
	setJDueProperty(it,"1/1/4501");it.Save();

  doWidgetRefresh(Ext.getCmp("w"+id),id);
	}

}

function removeTagJI(tg,id,iid)
{
//remove tag by x from jello inbox widget
var tg=tg.replace(new RegExp("~","g")," ");
var it=NSpace.GetItemFromID(id);
var tags=it.itemProperties.item(catProperty).Value;
var newval="";
var e=tags.replace(new RegExp(",","g"),";");
var tag=e.split(";");
for (var x=0;x<tag.length;x++)
{
if (trimLow(tg)!=trimLow(tag[x])){newval+=tag[x]+";";}
}
it.itemProperties.item(catProperty).Value=newval;
it.Save();
doWidgetRefresh(Ext.getCmp("w"+iid),iid);
}

function wFolSearch(id)
{
//quick search OL folder widget
//Settings:0=Outlook inbox folder ID

var ws=getWidgetSettings(id);
var folder=setAndCheckArcFolder(ws[0]);

  if(wFolSearch.arguments.length==5)
  {//set only title
setTimeout(function(){
try{Ext.getCmp("w"+id).setTitle(txtFind+" "+folder);
return;
}catch(e){}},30);
}

  if(wFolSearch.arguments.length==2)
  {//customize

  var war = [new Ext.form.TriggerField({fieldLabel:txtOLFolder,value:folder.FolderPath,width:200,onTriggerClick:function(e){wOutlookViewSelectFolder(true,true);},id:'folfolder'}),new Ext.form.Hidden({value:folder.EntryID,id:'folfolderID'})];
  return war;
  }


if (folder!=null)
{


var ret="new Ext.Panel({id:'wcfs00',frame:false,margins:'5 5 5 5',html:renderFolSearch("+id+",'"+ws[0]+"')})";

}
else
{
var ret="new Ext.Panel({id:'wci00',frame:false,margins:'5 5 5 5',bodyStyle:'padding:0px 5px 5px 5px',html:'<b>Folder does not exist</b>'})";
}

setTimeout(function(){
try{Ext.getCmp("w"+id).setTitle(txtFind+" "+folder);
}catch(e){}},30);

return ret;
}

function renderFolSearch(id,fol)
{
//render quick search OL folder widget
var olf=NSpace.GetFolderFromID(fol);
var ht=getWidgetHeight(id);
var ret="";
ret+="<P ALIGN=CENTER><b>"+txtFind+"</b>&nbsp;<input id='olswe"+id+"' onkeydown=WOLFind("+id+",'"+fol+"') title='Enter search string to find items of "+olf+"' style='width:75%;border:1px gray solid;font-size:10px;' type=text size=30>";
ret+="<br><div style='padding:3px 6px 3px 6px;overflow-y:scroll;height:"+(ht-40)+";width:99%' id='olswediv"+id+"'><b><i>"+olf.FolderPath+"</i></b><br>Folder Contains "+olf.Items.Count+" "+txtItems+"</div></P>";
return ret;
}

function WOLFind(id,fol)
{
//execute quick search for OL folder widget
if(event.keyCode==13)
{
event.returnValue=false;
var olf=NSpace.GetFolderFromID(fol);
var ret="<table width=98%>";
var fld="urn:schemas:httpmail:subject";
var fname="Subject";
var iimg="widget-add.gif";
if (olf.DefaultItemType==2){/*Contact*/fld="urn:schemas:contacts:cn";var fname="FullName";iimg="user.gif";}
if (olf.DefaultItemType==1){/*appointment*/iimg="appoint.gif";}
if (olf.DefaultItemType==4){/*journal*/iimg="appoint.gif";}
if (olf.DefaultItemType==0){/*mail*/iimg="inboxmail.gif";}
if (olf.DefaultItemType==5){/*note*/iimg="note.gif";}
if (olf.DefaultItemType==6){/*post*/iimg="post.gif";}
if (olf.DefaultItemType==3){/*task*/iimg="task.gif";}

var sv=document.getElementById("olswe"+id).value;
  var oje=olf.Items;
	var its=oje.Restrict("@SQL="+fld+" like '%"+sv+"%'");
	var ct=its.Count;
	if (ct>0)
	{
		for (var x=1;x<=ct;x++)
		{
		var i=its(x);
		ret+="<tr><td width=1><img src='img\\"+iimg+"'></td><td><span class=widgettext><a class=jellolinktop onclick=olItem('"+i.EntryID+"')>"+i.ItemProperties.Item(fname).Value+"</a></span></td></tr>";
		}
	}
	else{ret+="Nothing Found";}

document.getElementById("olswediv"+id).innerHTML=ret;
}
}