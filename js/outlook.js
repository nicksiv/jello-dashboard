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

var IsPreviewPaneVisible=false;

var inboxTable;
var jelloFolderTypes = [];

function connect()
{
ol=GetOutlookApplication();
if (ol!=false && typeof(ol)!="undefined"){conStatus="Outlook";progressBar();}
else{conStatus=null;}

if (conStatus==null)
{
try{progressBar();
    ol=new ActiveXObject("Outlook.Application");
    conStatus="Outlook ActiveX";
    progressBar();
    }catch(d){conStatus=null;}

}


if (conStatus==null)
{
//cannot connect with current config
document.write("<img src="+imgPath+"biglogo.gif><div style=background:#f0f0f0;border:1px solid black;><br><span class=jellolinktop style=font-size:250%;font-weight:bold;>Jello.Dashboard cannot run in this environment</span><br><ul style=font-size:200%;padding-left:20px;><li>Assign it as a home page for an Outlook folder<li>Run the jello5.hta application<li>Open it with Internet Explorer browser<li>Use firefox with <a href='https://addons.mozilla.org/en-US/firefox/addon/1419' class=jellolinktop>IETab</a> extension</ul><br></div>");
return false;
}



NSpace=ol.GetNameSpace("MAPI");

delItems=NSpace.GetDefaultFolder(3);
sentItems=NSpace.GetDefaultFolder(5);
journalItems=NSpace.GetDefaultFolder(11);
jelloFolderTypes.push(["journal", journalItems.DefaultItemType, journalItems.DefaultMessageClass]);
outItems=NSpace.GetDefaultFolder(4);
draftItems=NSpace.GetDefaultFolder(16);
taskItems=NSpace.GetDefaultFolder(13);
jelloFolderTypes.push(["task", taskItems.DefaultItemType, taskItems.DefaultMessageClass]);
contactItems=NSpace.GetDefaultFolder(10);
jelloFolderTypes.push(["contact", contactItems.DefaultItemType, contactItems.DefaultMessageClass]);
calendarItems=NSpace.GetDefaultFolder(9);
jelloFolderTypes.push(["calendar", calendarItems.DefaultItemType, calendarItems.DefaultMessageClass]);
inboxItems=NSpace.GetDefaultFolder(6);
jelloFolderTypes.push(["mail", inboxItems.DefaultItemType, inboxItems.DefaultMessageClass]);
// add a message type for meeting requests, which aren't default type
jelloFolderTypes.push(["mail", "IPM.Schedule.Meeting.Request",99]);
noteItems=NSpace.GetDefaultFolder(12);
OLversion=parseInt(ol.version);
status=txtStartSettings;





//hide preview pane
/*
try{
IsPreviewPaneVisible=ol.ActiveExplorer.IsPaneVisible(3);
ol.ActiveExplorer.ShowPane(3,false);
document.body.style.height=screen.height;
}
catch(e){}*/
return true;
}

function GetOutlookApplication(){
//**j5
try{
var	a=window.external.OutlookApplication;
return a;}catch(e){return false;}
}

// get a table object for a folder
// sort as specified and filter as specified
function getTable(folder,filter)
{
var table;
var ret = null;
var defCol = "CreationTime";
if( folder.DefaultItemType == 0)
	defCol = "ReceivedTime";
table = folder.GetTable(filter);
table.Columns.RemoveAll();
table.Columns.Add("EntryID");
table.Columns.Add(defCol);

//if( filter != "") ret= table.Restrict(filter);
//else ret = table;
//if( sort == true) ret.Sort(defCol);
table.Sort(defCol, true);
//return ret;
return table;
}

function getFolderTable(tab,names)
{
	var nm;
	var ncols = tab.Columns.Count;
	// first 2 are entryID and date field, drop rest of defaults
	for( var i=ncols;i > 2;i--)
		tab.Columns.Remove(i);
	// now add in the columns for the type of record
	// start at 2, since the first 2 are the entryID and default sort field and already present
	for(var i=2; i<names.length; i++)
	{

		if((nm = names[i]) == "")
			continue;
		tab.Columns.Add(nm);
	}
}

function getTableArray(tab, start,end)
{
	var tabArray = new Array();

	// reset table to beginning
	try{tab.MoveToStart();}catch(e){return;}
	// skip to the row we need for the array
	if( start > 0)
	{
		for( var i=0; i < start; i++)
			tab.GetNextRow();
	}
	// figure out the number of rows we need
	// get those rows and convert from vbarray to a flattened version of a jscript array
	var len = end-start;
	var vbarray = tab.GetArray(len);
	var ar = vbarray.toArray();
	// now conver the flattened array to a 2 dimensional array
	// one attribute per row, each row having that attribute for all items
	var cols = tab.Columns.Count;
	for( var i=0; i < cols; i++){
		var tmpArray = [];
		var st = i*len;
		var stp = (i+1)*len;
		for( var j =st; j < stp; j++)
			tmpArray[j-st]=ar[j];
		tabArray.push(tmpArray);
	}
	return tabArray;

}

// used to create a new folder and add to the folder list in tree
function olAddFolder(it,nd)
{
	addFolderTemp = it;
	addFolderNd = nd;
	Ext.Msg.prompt(txtAddFolder, txtAddFolderPrmpt,function(bt,fname){
		if( bt == "ok"){
			try{var t=NSpace.GetFolderFromID(addFolderTemp);
				t.Folders.Add(fname);
				folderListObjs.splice(0,folderListObjs.length);
			//	var par = addFolderNd.parentNode;
				addFolderNd.collapse();
				addFolderNd.removeAll();
				}
			catch(e){return;}
		}
	});

}
function olDelFolder(it,nd)
{
	try{
		var t=NSpace.GetFolderFromID(it);
		t.Delete();
		folderListObjs.splice(0,folderListObjs.length);
		var par = nd.parentNode;
		par.collapse();
		par.removeAll();
	}catch(e){}


}
// convert a message type default to a classid
function mtypeToClass(mtype)
{
	var ret = null;
	for( var i=0; i <jelloFolderTypes.length; i++)
	{
		// if the type match the folder type
		// OL can have a type that start with IPM.Note but then adds .signed.trusted, so
		// we use indexof
		if( mtype.indexOf(jelloFolderTypes[i][2]) != -1){
			switch(jelloFolderTypes[i][1]){
			case 0:
				return 43;
			case 1:
				return 26;
			case 2:
				return 40;
			case 3:
				return 48;
			case 4:
				return 42;
			case 5:
				return 44;
			case 6:
				return 45;
			case 7:
				return 69;
			case 99:
				return 53;
			default:
				break;
			}
		}
	}
	return ret;
}


function olItem(id)
{
//j5
//open outlook item
try{
	var t=NSpace.GetItemFromID(id);
}
catch (e){
alert(txtOItemDeleted);
}

if (typeof(t) == 'undefined') {
	alert(txtOItemDeleted);
	return;
} else {
  t.Display();
  }
}

function isFolder(id)
{
try
{var i=NSpace.GetFolderFromID(id); return true;}
catch(e)
{return false;}
}

function addOLItemCategory(olitem,tagname)
{

var cats=olitem.itemProperties.item(catProperty).Value;
cats=cats.replace(new RegExp(",","g"),";");
var c=cats.split(";");
var flag=false;
	for (var x=0;x<c.length;x++)
	{
		if (trimLow(c[x])==trimLow(tagname)){flag=true;}
	}

if (flag==false)
{olitem.itemProperties.item(catProperty).Value+=";"+tagname;
olitem.Save();}
}

function iconBasedOnClass(c,cdef)
{
//return an icon for items based on item's class
var ic=icTask;
if (notEmpty(cdef)){ic=cdef;}

			if (c==26 || c==53){ic=icApp;}
			if (c==40){ic=icContact;}
			if (c==43){ic=icMail;}
			if (c==44){ic=icNote;}
			if (c==42){ic=icJournal;}
			if (c==45){ic=icPost;}
			if (c==0){ic="mng.gif";}
return ic;
}

function OLDeleteItem(id){
//j5
//delete an outlook item (move to trash)
showOLViewCtl(false);
// in the master list view you can have 2 recs for an item
// so the lookup can throw exception, put a try around it all
try{
	var d=NSpace.GetItemFromID(id);
	//for tasks remove its reference from linked items...
	if (d.Class==48)
	{
	var tLinks="";
	try{tLinks=d.UserProperties.Item("OLID").value;}catch(e){}
	if (notEmpty(tLinks))
	{
		var tl=tLinks.split(";");
		if (tl.length>0)
		{
			for (var x=0;x<tl.length;x++)
			{
				try{
				var it=NSpace.GetItemFromID(tl[x]);
				try{it.UserProperties.Item("OLID").Value="";it.Save();}catch(e){}

			  if (jello.removeLinked==0 || jello.removeLinked=="0")
			  {
			  var cmsg=txtConfAttDl.replace("%1",it).replace("%2",it.Parent.FolderPath);
			  //choice=confirm(cmsg);

			  var buts={yes:txtDelete,no:txtMoveInfo,cancel:txtCancel};

			  Ext.Msg.show({
				   //title:txtAppointment,
				   msg: cmsg,
				   buttons: buts,
				   fn: function(b,t){OLDelOrMove(b,t,it);},
				   animEl: 'elId',
				   icon: Ext.MessageBox.QUESTION
				});
			  }else if (jello.removeLinked==1 || jello.removeLinked=="1"){
				try{it.Delete();}catch(e){alert("This item could not be deleted\nYou should use Outlook to do that!");}
			  }



			  }	catch(e){}
			}
		}
	}
	}
	d.Delete();
}catch (e){alert("Could not perform selected action on Outlook item.");return true;};
showOLViewCtl(true);
updateInboxCount();
updateTheLatestThing();
}

function OLDelOrMove(b,rc,it)
{
	// yes means delete
	// no means move
	// cancel means do nothing
	if( b == "yes")
	{
	try{it.Delete();}catch(e){}
	}else if( b == "no"){
		var t=NSpace.PickFolder();
		if( t == null)
			return;
		if (t.DefaultItemType==0){

			try{it.Move(t);}catch(e){alert(txtMsgNoMoveErr);return;}
		}else
			try{it.Copy(t);}catch(e){alert(txtMsgNoMoveErr);return;}
	}

showOLViewCtl(true);
}


function OLFolderOpenNewWindow(id)
{
//open outlook folder in new window
if (!notEmpty(id)){id=jello.archiveFolder;}
try{var fol=NSpace.GetFolderFromID(id);}catch(e){alert(txtInvalid);}
var exp=fol.GetExplorer();
exp.Display();
}

function editOLFolder(id)
{
//Outlook folder properties edit

try{var fol=NSpace.GetFolderFromID(id);}catch(e){alert(txtInvalid);}
var finfo="";
var r=null;

try{
var ix=olFolderStore.find("eid",new RegExp("^"+id+"$"));
r=olFolderStore.getAt(ix);
}catch(e){r=null;}

finfo="<p align=center><b>"+fol.Items.Count+"</b> "+txtItemItems+" (<b>"+fol.UnreadItemCount+"</b> "+txtUnread+").<br>";
if (id==jello.inboxFolder){finfo+="<span class=tagprj>Jello "+txtJInbFol+"</b>&nbsp;";}
if (id==jello.actionFolder){finfo+="<span class=tagprj>Jello "+txtJActFol+"</b>&nbsp;";}
if (id==jello.archiveFolder){finfo+="<span class=tagprj>Jello "+txtJArcFol+"&nbsp;";}
if (id==jello.calendarFolder){finfo+="<span class=tagprj>Jello "+txtJTicFol+"&nbsp;";}
if (id==jelloSettingsFolder){finfo+="<span class=tagprj><b>Jello "+txtJSetFol+"</b>&nbsp;";}
if (shortcutExists("outlookView('"+id+"')")){finfo+="<span class=tagnext>"+txtMsgShocCre;}
finfo+="</P>";

var olvws=new Array();
  for (var x=1;x<=fol.Views.Count;x++)
  {olvws.push(fol.Views.Item(x).Name);}

if (r==null)
{
var ottl=fol;
var oax=false;
var owin=false;
var ocount=false;
var adolview=fol.Views.Item(1).Name;
}
else
{
var ottl=r.get("fname");
var oax=r.get("activex");
var owin=r.get("newwindow");
var ocount=r.get("counter");
var adolview=r.get("defolview");
}
if (typeof(adolview)=="object"){adolview=adolview.Name;}
if (!notEmpty(adolview)){adolview=fol.Views.Item(1).Name;}
var showdolview=false;
oview=txtOLJView1;
if (oax){oview=txtOLJView2;}
if (owin){oview=txtOLJView3;}
if (oview==txtOLJView2){showdolview=true;}

try{var odesc=fol.Description;}catch(e){var odesc="";}
try{var oentryid=fol.EntryID;}catch(e){var oentryid="None";}

var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: fol.FolderPath,
        height:380,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'olform',
        iconCls:'olformicon',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',

        items: [{
                fieldLabel: txtFolName,
                id: 'oname',
                value:ottl,
                disabled:false,
                tabIndex:0,
                allowBlank:false
            },{
                fieldLabel: 'Outlook ID',
                id: 'olid',
                value:oentryid,
				readOnly:true,
				style:'font-size:9px;background:none;border:none;'
            },
                new Ext.form.TextArea({
                fieldLabel: txtNotes,
                id: 'onotes',
                height:100,
                value:odesc
            }),
                    new Ext.form.ComboBox({
                fieldLabel: txtView,
                id: 'iview',
                store:new Array(txtOLJView1,txtOLJView2,txtOLJView3),
                hideTrigger:false,
                typeAhead:false,
                editable:false,
                triggerAction:'all',
                value:oview,
                listeners:{select: function(cb,rec,idx){
                var cva=Ext.getCmp("iview").getValue();
                var showov=false;
                if (cva==txtOLJView2){showov=true;}
                if (showov){Ext.getCmp("olviewctl").show();}else{Ext.getCmp("olviewctl").hide();}
         		   //Outlook 2010 can have native views only in IE or hta mode
                if (OLversion>13 && conStatus=="Outlook" && cva==txtOLJView2)
               {alert("You will not be able to view in that mode into Outlook with version 2010. Use standalone or IE version of the application.");}

                }},
                mode:'local'
                }),
                new Ext.form.ComboBox({
                fieldLabel: txtOLView,
                id: 'olviewctl',
                store:olvws,
                hideTrigger:false,
                typeAhead:false,
                editable:false,
                triggerAction:'all',
                hidden:!showdolview,
                value:adolview,
                mode:'local'
                }),
                new Ext.form.Checkbox({
                fieldLabel: txtCounter,
                id: 'ocount',
				checked:ocount
            }),
            new Ext.form.Label({
				html:finfo
            }),
                        new Ext.form.Hidden({
				id:'oid',
				value:id
            })
        ]

    });



   var win = new Ext.Window({
        title: txtOLFolderProps,
        width: 480,
        height:400,
        id:'theolform',
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
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text:txtSave,
            id:'sbut',
            tooltip:txtSave+' [F2]',
            listeners:{
            click: function(b,e){
            saveOLForm(b,win);
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
  setTimeout(function(){Ext.getCmp("oname").focus(true,100);},1);

//shortcut keys
var map = new Ext.KeyMap('theolform', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});

var map2 = new Ext.KeyMap('theolform', {
    key: Ext.EventObject.F2,
    fn: function(){
    try{
    saveOLForm(null,win);
    }catch(e){}},
    scope: this});

}

function saveOLForm(b,win)
{
//save Outlook properties form
try{var id=Ext.getCmp("oid").getValue();}catch(e){alert(txtInvalid);return;}
try{var fol=NSpace.GetFolderFromID(id);}catch(e){alert(txtInvalid);return;}
var tname=Ext.getCmp("oname").getValue();
var tnotes=Ext.getCmp("onotes").getValue();
var oview=Ext.getCmp("iview").getValue();
var olviewCtl=Ext.getCmp("olviewctl").getValue();
var tax=false;
var twin=false;
if (oview==txtOLJView2){tax=true;}
if (oview==txtOLJView3){twin=true;}
var tcou=Ext.getCmp("ocount").getValue();

try{
var ix=olFolderStore.find("eid",new RegExp("^"+id+"$"));
var r=olFolderStore.getAt(ix);}catch(e){var r=null;}
var tid=id;

  if (r==null)
  {
  //no existing record
var tr=new olFolderRecord({
eid:tid,
fname:tname,
activex:tax,
newwindow:twin,
counter:tcou,
defolview:olviewCtl
});
olFolderStore.add(tr);
  }
  else
  {
  r.beginEdit();
  r.set("fname",tname);
  r.set("activex",tax);
  r.set("newwindow",twin);
  r.set("counter",tcou);
  r.set("defolview",olviewCtl);
  r.endEdit();
  }

  fol.Description=tnotes;

var msg=txtMsgOLFldIfUpd+"<br>";

syncStore(olFolderStore,"jello.outlookFolders");

//change shortcut's name if exists
var ix=shortcutStore.find("cmd","outlookView('"+id+"')");
var r=shortcutStore.getAt(ix);
var cnt="";
if (tcou==true){cnt=getOLFolderCounters(id);}

  if (r!=null)
  {
  r.beginEdit();
  r.set("shortcut",tname);
  r.endEdit();
  msg+="Shortcut updated";
  //update sortcut node

  var nd=Ext.getCmp("tree").getNodeById("s:"+r.get("id"));
  nd.setText("<span class=customOLfolder>"+tname+"</span>"+cnt);
  }
syncStore(shortcutStore,"jello.shortcuts");

//save settings
jese.saveCurrent();
//update olfolder node
var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();

	try{var trfid=nd.id;
	if (trfid=="o:"+id){nd.setText("<span class=customOLfolder>"+tname+"</span>"+cnt);}
	}
	catch(e){}

Ext.info.msg(txtEdOLFolder, msg);
try{win.destroy();}catch(e){}
}

function jelloOLFName(thisfold,id)
{
//return the jello custom outlook folder name if any
var rv=thisfold;
//try{var oc=olFolderStore.getCount();}catch(e){return rv;}
var ix=olFolderStore.find("eid",new RegExp("^"+id+"$"));
var r=olFolderStore.getAt(ix);


  if (r!=null)
  {
  var cnt="";
  if(r.get("counter")==true){cnt=getOLFolderCounters(id);}
  rv="<span class=customOLfolder>"+r.get("fname")+"</span>"+cnt;
  }

return rv;
}

function getOLFolderCounters(id)
{
if (id.substr(0,7)=="outlook"){id=id.replace("outlookView('","");id=id.replace("')","");}
  try{
  var f=NSpace.GetFolderFromID(id);
  }catch(e){return "";}
  return "&nbsp;("+f.UnReadItemCount+"/"+f.Items.Count+")";
}

function olviewCtl(fid,olname)
{
//build outlook view control
//useless at the moment
var ret="";
var ff=NSpace.GetFolderFromID(fid);
var fff=ff.FolderPath;
if (!notEmpty(olname)){olname="olnative";}
//var tp=Ext.getCmp("west-panel").getPosition();
//var hgt=Ext.getCmp("west-panel").getInnerHeight()-tp[1];
var panelH=Ext.getBody().getViewSize().height;
var prp=0;
try{if (ff.DefaultItemType==1 && OLversion>11 && OLversion<14){prp=80;Ext.getCmp("olview").setHeight(Ext.getCmp("olview").getHeight()-65);}}catch(e){}
//prp=jello.actionPreviewHeight;
var hgt=(panelH-prp);
      var marg=0;
      if (OLversion>=11 && jello.addSpaceOlView==true){marg=OL2007VwMargin;}
ret="<div id="+olname+"space style=background:#D3E1F1;height:100px;display:none;></div><OBJECT id="+olname+" style='margin-top:"+marg+"px;' CLASSID='clsid:0006F063-0000-0000-C000-000000000046' height='"+((hgt-170)-marg)+"'' width=100%><PARAM name=Folder value='" + fff + "'>";
ret+="<PARAM name=DeferUpdate value=false></OBJECT>";

return ret;
}

//unused. stopped tracking outlook preview pane
function restoreOutlookDefaults()
{
//restore outlook view defaults
if (jello.hideOLPreviewPane==false || jello.hideOLPreviewPane==0 || jello.hideOLPreviewPane=="0"){return;}
try{
ol.ActiveExplorer.ShowPane(3,IsPreviewPaneVisible);
}
catch(e){}

}

function pOLInbox(folder)
{
//Outlook folder view in Jello
getTagsArray();
var isOL2010=!OLversion<14;
initScreen(false,"pOLInbox("+folder+")");
var isActiveX=false;
if (conStatus=="Outlook ActiveX"){isActiveX=true;}

var oViews=new Array();
	for (var x=1;x<=folder.Views.Count;x++)
	{
	try{oViews.push(folder.Views.Item(x).Name);}catch(e){}
	}

 tbar2 = new Ext.Toolbar({
 id:'gridfooter',
items:[
new Ext.form.Label({
				html:"0 " + txtItems,
				id:'acounter'
				}),
        '-',
        {
        iconCls:'olpreview',
        text:txtOPreview,
        id:'oprevw',
        hidden:isActiveX,
        tooltip: txtOLPwPaneTgl,
			listeners:{click: function(b,e){
			toggleOLViewPreview();}}
        },'-',
        new Ext.form.ComboBox({
                id: 'iviews',
                store:oViews,
                hideTrigger:false,
                typeAhead:false,
                triggerAction:'all',
                listWidth:300,
                emptyText:txtOLView,
                mode:'local',
                maxHeight:100,
                listeners:{
            select: function(cb,rec,idx){
            setOLView();
            },
            expand: function(c){showOLViewCtl(false);},
			collapse: function(c){showOLViewCtl(true);}
			}
            }),'-',{
        text:txtFilter,
	    id:'ifilter',
	    tooltip:txtFltrSenderIfo,
	    enableToggle:true,
	    hidden:isOL2010,
		listeners:{
            click: function(b,e){filterOLView();}
            }
            },'-',{
        text:txtProps,
	    id:'ifolder',
		listeners:{
            click: function(b,e){editOLFolder(folder.EntryID);}
            }
            }


	] });

var ttb=msgToolbar(true);
//add olview due date button
        ttb.add({
    cls:'x-btn-icon',
    icon: 'img\\calendar.gif',
    tooltip:txtDueForInfo,
    id:'odate',
    listeners:{click: function(b,e){
            var dd=new Date();
            var ddf=dd.format('m/j/Y');
            if(jello.dateFormat == 0 || jello.dateFormat == "0"){ddf=dd.format('j/m/Y');}
            if(jello.dateFormat == 2 || jello.dateFormat == "2"){ddf=dd.format('Y/m/d');}
			var dp=prompt(txtMsgOLVDate,ddf);
			if (!isDate(dp)){return;}
			var date=new Date(dp);
            var dt=date.format('m/j/Y');
            if(jello.dateFormat == 0 || jello.dateFormat == "0"){dt=date.format('j/m/Y');}
            if(jello.dateFormat == 2 || jello.dateFormat == "2"){dt=date.format('Y/m/d');}
            inboxAction(null,'due',dt);
    }
    }
});


var grid=new Ext.Panel({
         frame:true,
         border:false,
         id:'olview',
         tbar:ttb,
         bbar:tbar2,
         height:(panelHeight-100),
         autoWidth:true
});


var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.doLayout();
grid.update(olviewCtl(folder.EntryID));
grid.doLayout();
ppnl.setTitle("<span><img src=img\\"+icFolder+" style=float:left;> "+folder.FolderPath);
//Ext.getCmp("ipit").hide();Ext.getCmp("init").hide();
Ext.getCmp("popdate").hide();

//in OL 2010 there is no filter property of OLViewControl Using Filters collection instead
if (OLversion<14)
{if (notEmpty(olnative.Filter)){olnative.Filter="";}}

if (OLversion<=10){Ext.getCmp("oprevw").hide();}
updateOLViewCounter();
setThisOLView(folder.EntryID);
//status=txtReady;
}

function showOLViewCtl(show)
{//show/hide the outlook view control
var ss="hidden";
if (show){ss="visible";}
try{olnative.style.visibility=ss;}catch(e){}
//hide/show widgets
  for (var x=1;x<20;x++)
  {
  try{var ov=document.getElementById("olv:"+x);ov.style.visibility=ss;}catch(e){}
  }
}

function lowerOLViewCtl(lower,idname)
{
if (jello.addSpaceOlView==0){return;}
var olv="olnativespace";
if (notEmpty(idname)){olv=idname+"space";}
var ov=document.getElementById(olv);
  if (lower)
  {
  try{
  ov.style.display="block";
  }catch(e){}
  }
  else
  {
  try{
  ov.style.display="none";
  }catch(e){}
  }
}

function toggleOLViewPreview(force)
{
//toggle outlook view preview on/off
  try
  {
 var isVis=ol.ActiveExplorer.IsPaneVisible(3);
  if (force!=null){ol.ActiveExplorer.ShowPane(3,force);return;}
   ol.ActiveExplorer.ShowPane(3,!isVis);
  }catch(e){}
}

function setOLView()
{
//change OL folder viewer View
var c=Ext.getCmp("iviews").getValue();
try{olnative.View=c;}catch(e){}
}


function filterOLView()
{//filter the outlook view control
var press=Ext.getCmp("ifilter").pressed;
if (press==false){olnative.Filter="";olnative.SaveView(olnative.View);return;}
var ft=prompt(txtFltOLView,"");
  if (notEmpty(ft))
  {

      olnative.Filter="urn:schemas:httpmail:subject LIKE '%"+ft+"%'";

  }
}

function setThisOLView(id)
{
//use saved view for outlook view control from outlook properties settings
try{
var ix=olFolderStore.find("eid",new RegExp("^"+id+"$"));
r=olFolderStore.getAt(ix);
}catch(e){return;}
var vw="";
if (r==null){vw=NSpace.GetFolderFromID(id).Views.Item(1).Name;}
else{var vw=r.get("defolview");}
if (typeof(vw)=="object"){vw=vw.Name;}
  if (notEmpty(vw)){
  try{olnative.View=vw;}catch(e){}
  }
}

// get a delomited section of the body of an item
function getOLBodySection(it,delim)
{
	var jn = "";
	try{var oln = it.Body;}catch(e){return null;}
	var jnStart = oln.indexOf(delim);
	if( jnStart == -1)
		return null;
	// if start. find end
	var jnEnd = oln.indexOf(delim, jnStart + delim.length);
		// if an end, we have notes
	if( jnEnd == -1)
		return null;
	jn  = oln.substring(jnStart+delim.length, jnEnd);
	jn = jn.replace(/\r*/g,"");
	jn = jn.replace(/^\n+/g,"");
	return jn;
}
// add/update a delimited section to the body
function setOLBodySection(it,delim,val)
{
		var addnl ="";
		if( val.charAt(val.length-1) != '\n')
			addnl = "\n";
		var oln=it.Body;
		// find jello notes start tag
		var jnStart = oln.indexOf(delim);
		// note there yet, put at front
		if( jnStart == -1){
			it.Body = delim + "\n" + val +
				addnl + delim + "\n" + oln;
		}else{
			var jnEnd = oln.indexOf(delim, jnStart + delim.length);
			// if there is start/end replace middle
			if( jnEnd != -1){
				var notes = val;
				notes = notes.replace(/^[\n\r]*/,"");
				notes = notes.replace(/[\n\r]*$/,"");
				// if not start at beginning
				//.put other test as start
				var newBody = "";
				if( jnStart != 0)
					newBody = oln.substring(0,jnStart);
				// then the n delim, nites, delim
				// and stuff from body after trailing delim
				newBody += delim + "\n" + notes + addnl + delim + "\n"+ oln.substring(jnEnd+delim.length);
				it.Body = newBody;

			}else{
				// no mathcing end, delete delim
				// and place new notes at start
				it.Body = delim + "\n" + val +
					addnl + delim + "\n"+ oln.substring(0,jnStart) +
					oln.substring(jnStart+delim.length);

			}
		}
		if(it.Body.charAt(it.Body.length-1) != '\n')
			it.Body += '\n';

}
// take body text and remove delimited section
function cleanJelloSectionFromOLBody(body,delim)
{
	var ret=body;
	// find jello notes start tag
	var jnStart = body.indexOf(delim);
	// if start. find end
	if( jnStart != -1){
		var olns = body.substring(0,jnStart);
		var jnEnd = body.indexOf(delim, jnStart + delim.length);
		// if an end, we have section we seek
		if( jnEnd != -1)
			ret  = olns + body.substring(jnEnd+ delim.length);
	}
	return ret;
}

function uninst()
{
//Unset outlook folder to show the dashboard
if (confirm("Unset Jello Dashboard from folder?")==false){return;}
ol.ActiveExplorer.CurrentFolder.WebViewOn=false;
alert(txtMsgUninstall);
}

function getJPriorityProperty(it){
//get custom priority field value. Create if not exists
try{

var jn="";
try{
jn=it.UserProperties.Item("jpri").Value;}
catch(e)
{
it.UserProperties.Add("jpri",3,true);
if (it.Importance==2){it.UserProperties.Item("jpri").Value=1;it.Save();jn=1;}
}
if (jn=="" || jn==null){jn=99;}
}catch(ee){jn=99;}
return jn;

}

Ext.override( Ext.form.ComboBox, {
    anyMatch: true,
    /* Set to true to match case-insenstively: */
    queryIgnoreCase: false,
    /* Set to true to only match on word start (beginning of string,
       after whitespace or quotation marks: */
    matchWordStart: false,
    doQuery : function(q, forceAll){
        if(q === undefined || q === null){
            q = '';
        }
        var qe = {
            query: q,
            forceAll: forceAll,
            combo: this,
            cancel:false
        };
        if(this.fireEvent('beforequery', qe)===false || qe.cancel){
            return false;
        }
        q = qe.query;
        forceAll = qe.forceAll;
        if(forceAll === true || (q.length >= this.minChars)){
            if(this.lastQuery !== q){
                this.lastQuery = q;
                if(this.mode == 'local'){
                    this.selectedIndex = -1;
                    if(forceAll){
                        this.store.clearFilter();
                    }else{
                        var qr;
                       if( this.matchWordStart ){
                            if( this.queryIgnoreCase ){
                                qr = new RegExp( '\\b' + q, "i" );
                            } else {
                                qr = new RegExp( '\\b' + q );
                            }
                        } else {
                            qr = q;
                        }
                        this.store.filter(this.displayField, qr, this.anyMatch, this.queryIgnoreCase);
                    }
                    this.onLoad();
                }else{
                    this.store.baseParams[this.queryParam] = q;
                    this.store.load({
                        params: this.getParams(q)
                    });
                    this.expand();
                }
            }else{
                this.selectedIndex = -1;
                this.onLoad();
            }
        }
    }
    ,onTypeAhead: function() {
        var nodes = this.view.getNodes();
        for( var i = 0; i < nodes.length; i++ )
        {
            var n = nodes[i];
            var d = this.view.getRecord( n ).data;
            var re;
            var ptrn = '(.*?)(' + (this.matchWordStart ? '\\b' : '') + this.getRawValue() + ')(.*)';

            if( this.queryIgnoreCase ){
                re = new RegExp( ptrn, "i" );
            } else {
                re = new RegExp( ptrn );
            }
            var h = d[this.displayField];
            h = h.replace( re, '$1<span class="mark-combo-match">$2</span>$3' );
            n.innerHTML = h;
        }
    }
});

var folderListObjs= new Array();
var tagListObjs = new Array();
var fnameStore;
function buildFolderList(refresh,tagsToo)
{
  
	if(refresh)
		folderListObjs.splice(0,folderListObjs.length);
	if( folderListObjs.length == 0)
	{
		// first time, so build the list
		var top = enumerateFoldersInStores();
		// get the root folder name
		var rootFolderName = NSpace.GetFolderFromID(jello.inboxFolder).Parent.FolderPath;
		rootFolderName = rootFolderName.slice(2);
		for(var i=0; i < top.length; i++){
			if( top[i].name == rootFolderName){
				var xf = new olFolderObject();
				xf.icon = top[i].icon;
				xf.id = top[i].id;
				xf.name = top[i].name;
				top.splice(i,1);
				top.splice(0,0,xf);
				break;
			}
		}

		var mbox = null;
		Ext.MessageBox.progress("Build Folder List","Processing Folders...","Starting...");

		for( var i=0; i < top.length; i++)
		{
			if( top[i].name.indexOf("Public") == 0)
				continue;
			try{top[i].name = NSpace.GetFolderFromID(top[i].id);} catch(e){continue;}
			Ext.MessageBox.updateProgress(i/top.length, "PST: " + top[i].name);
			buildOLFolderList(top[i],folderListObjs);
		}
	}
		 if (tagsToo==true){
		  //tags
		  var tgcnt=tagStore.getCount();
		  for (var x=0;x<tgcnt;x++)
			{
			var r=tagStore.getAt(x);
				var tn=r.get("tag");
		    var tg = new Array("-"+r.get("id"), "\\\\"+txtJObject+"\\"+tn.toLowerCase());
		  	tagListObjs.push(tg);
			}
	}
	var allObjs = new Array();

	if( tagsToo)allObjs = folderListObjs.concat(tagListObjs);
	else allObjs = folderListObjs.slice(0);
	for(var i=0; i < allObjs.length; i++)
		allObjs[i][2] = i;
	if( Ext.MessageBox.isVisible()) Ext.MessageBox.hide();
		fnameStore = new Ext.data.ArrayStore({
		fields: ['id','folderPath', {name:'xsort',type: 'int'}],
		sortInfo: {field: 'xsort', direction: 'ASC'},
		data : allObjs
		});

}

function goToFolder(tagsToo)
{
  var fref = outlookView;
	var captn=txtGoToFolder;
	Ext.MessageBox.progress("Go to Folder","Processing Folders...","Starting...");
	if (tagsToo==true){captn=txtGoto;}
	selFolder(captn, fref, null,tagsToo);
	Ext.MessageBox.hide();
}
function selFolder(boxTitle, funct, xdata,tagsToo)
{
		
		  
      if (typeof(fnameStore)=="undefined")
      {
      Ext.MessageBox.progress(boxTitle,"Processing Folders...","Starting...");
      buildFolderList(false,tagsToo);
      Ext.MessageBox.hide();
      }
      
		var cinfo=txtGotoInfoOL;
		if (tagsToo==true){cinfo=txtGotoInfo;}
		var combo = new Ext.form.ComboBox({
			id: "selFoldCombo",
			//title: "\nMessage(s) to Move/Copy\n",
			fieldLabel: txtFolder,
			//autoHeight: true,
			store: fnameStore,
			anyMatch: true,
			matchWordStart: false,
			queryIgnoreCase: true,
			displayField:'folderPath',
			typeAhead: true,
			mode: 'local',
			emptyText:txtGotoInfo+'...',
			selectOnFocus:true,
			resizable: true,
			//minListWidth: 110,
			anchor: "95% 10%",
			folData: xdata,
			folFunc: funct,
			folId: "",
			forceSelection: true,
			allowBlank: (xdata == null?false:true),
			listeners : {
				select : function(combo,record,idx){
					var id = record.get("id");
					Ext.getCmp("selFoldCombo").folId = id;
					//Ext.getCmp('selFolderOK').focus();
					(function(){ Ext.getCmp('selFolderOK').focus();}).defer(100);
			}
			}

		});
		var messageMark= [
           '<div onclick="openInboxItem(\'{entryID}\')">',
           '<tpl if="typeof(values.to) != \'undefined\'">',
           'To: {to}\n<br>',
            '</tpl>',
            'From: {sender}\n<br>',
    		txtSubject+ ':{subject}\n<br>',
			'</div><hr>'
       ];
       var thistemp = new Ext.XTemplate(messageMark);
       var iArray = new Array();

       if( xdata!=null){

       	var tArea = new Ext.Panel({
			id: "folderLines",
			anchor: "95% 60%",
			resizable: true,
			readOnly: true,
			//autoWidth: true,
			tplWriteMode: "overwrite",
			autoHeight: true,
			items: [
			{
			html:"<font style='font-size: " + jello.mailPreviewFontSize +"px'></font>",
			id:'selmtext',
			height: 280,
			autoWidth:true,
			border:true,
			margins: {bottom:10},
			autoScroll:true,
			bodyStyle: {
			background: '#ffffff',
			padding: '0px'
			},
			layout:'anchor'
			}]


       	});
       	iArray.push(tArea);
		// this next line is length of chars in each item, used with column below to get
		// things to fit correctly
		var chrs = txtTask.length+txtReply.length+txtFwd.length+txtAppointment.length+
			txtDelete.length+txtCopy.length+6;
		if( xdata.length == 1){
			var cbox = new Ext.form.CheckboxGroup({
				id: 'checkActs',
				xtype: 'checkboxgroup',
				columns: [(txtTask.length+1)/chrs,(txtReply.length+1)/chrs,
					(txtFwd.length+1)/chrs,(txtAppointment.length+1)/chrs,(txtCopy.length+1)/chrs,(txtDelete.length+1)/chrs],
				items: [
					{boxLabel: txtTask, name: 'folderMakeTask', id:"ckTask"},
					{boxLabel: txtReply, name: 'folderDoReply', id: "ckRepl"},
					{boxLabel: txtFwd, name: 'folderForward', id: "ckFwd"},
					{boxLabel: txtAppointment, name: 'folderDoAppt', id: "ckAppt"},
					{boxLabel: txtCopy, name: 'folderDoCopy', id: "ckCopy"},
					{boxLabel: txtDelete, name: 'folderDoDelete', id: "ckDel"}
				],
				listeners: {
					change: function(xobj, itms){
						for(var i=0; i < itms.length; i++)
						{
							if( itms[i].getId() == "ckTask" )
							{
								var elm = Ext.getCmp('mtags');
								if( itms[i].getValue() == true)
									elm.allowBlank = false;
								else
									elm.allowBlank = true;
								elm.render();
							}
						}
					}
				}

			});
			iArray.push(cbox);


		}

       	}
       	iArray.push(combo);
      	if( xdata != null && xdata.length == 1){
	// this code adds the categoty select
			var tgs="";var ftagList="";
			try{var it = NSpace.GetItemFromID(xdata[0][0]);}catch(e){return;}
			tgs=tagList('mformtagdisplay',null,it.itemProperties.item(catProperty).Value);
			ftagList=it.itemProperties.item(catProperty).Value;
			if(tgs==""){tgs=" ";}
			var cbox2 = new Ext.form.ComboBox({
				fieldLabel: txtTags,
				id: 'mtags',
				store:globalTags,
				hideTrigger:false,
				typeAhead:true,
				mode:'local',
				listeners:{
				blur:function(){var mx = Ext.getCmp("selFoldCombo");quickAddTag("mtags", "mtagfield",'mformtagdisplay',mx.folData[0][0]);},
				specialkey: function(f, e){
				if(e.getKey()==e.ENTER){
				userTriggered = true;
				e.stopEvent();
				f.el.blur();
				quickAddTag("mtags", "mtagfield",'mformtagdisplay',Ext.getCmp("selFoldCombo").folData[0][0]);}},

				select: function(cb,rec,idx){
				quickAddTag("mtags", "mtagfield",'mformtagdisplay',Ext.getCmp("selFoldCombo").folData[0][0]);
				}
				}

       		});
			iArray.push(cbox2);
			var catdisp = new Ext.form.Label({
			html:tgs[0],
			height:40,
			autoWidth: true,
			//cls:'formtaglist',
			id:'mformtagdisplay'
			});
			iArray.push(catdisp);

			var cats = new Ext.form.Hidden({
			id:'mtagfield',
			value:ftagList
			});
			iArray.push(cats);
       	}
	var gotoF;
	if( xdata != null)
		gotoF = new Ext.form.FormPanel({
		labelWidth: 50,// label settings here cascade unless overridden
        frame:true,
		//layout: 'anchor',
		resizeable: true,
		autoScroll: true,
        title: "",
        width:580,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        id:'mvform',
        buttonAlign:'center',
        defaults: {width: 400},
        defaultType: 'textfield',
		items: iArray
	});



	   var win = new Ext.Window({
        title: boxTitle,
        width: (xdata == null?400: 600),
        //height: 500,
        anchorSize: {width:300},
        //autoHeight: true,
        id:'selectfold',
        resizable:true,
        //draggable:true,
        layout: 'anchor',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: (xdata == null?iArray:gotoF),
        modal:true,
        listeners: {destroy: function(t){showOLViewCtl(true);},
		specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();}}
		},
        buttons: [
			{text: (xdata==null?txtGoToFolder:txtExecuteActions),
			id: 'selFolderOK',
            listeners:{
            click: function(b,e){
            	var c = Ext.getCmp("selFoldCombo");
            	var id = c.folId;
		        if (id.substr(0,1)!="-")
				{
				var f = c.folFunc;
				var d = c.folData;
				var facts = Ext.getCmp("checkActs");
				//var cked = null;
				//var rked = null;
				//var aked = null;
				var xacts = {task:null,reply:null,appt:null, fwd: null, del: null, copy: null, cats: ""};
				if( typeof(facts) != "undefined") {
					var factvals = facts.getValue();
					for(var i=0; i< factvals.length; i++)
					{
						var lbl = factvals[i].getId();
						switch(lbl){
						case "ckTask": xacts.task= factvals[i].getValue();break;
						case "ckRepl": xacts.reply = factvals[i].getValue();break;
						case "ckAppt": xacts.appt = factvals[i].getValue();break;
						case "ckFwd": xacts.fwd = factvals[i].getValue();break;
						case "ckDel": xacts.del = factvals[i].getValue();break;
						case "ckCopy": xacts.copy = factvals[i].getValue();break;
						}
					}
				}
				var mf = null;
				if(( mf = Ext.getCmp('mtagfield')) != null && mf != "undefined") xacts.cats = mf.getValue();
				if( xacts.task == true && xacts.cats ==""){
					Ext.getCmp('mtags').render();
					Ext.getCmp('mtags').focus(null,100);
					return false;
				}

				if( d == null)
				{	f(id);}
				else
				{	//f(id,d,cked, rked, aked);
					f(id,d,xacts);
				}

		 		} else{
		         //open tag
		         var tgid=parseInt(id);
		         tgid=tgid+((-tgid)*2);

		          showContext(tgid,1);

		         }
				 setTimeout(function(){
				        var w = Ext.getCmp("selectfold");
								w.close();
				        },6);

			}
            }},
			{text: "Show Folder Items",
			id: 'selFolderShow',
			hidden: (xdata == null?"true":"false"),
            listeners:{
            click: function(b,e){
            	var c = Ext.getCmp("selFoldCombo");
            	var id = c.folId;
		        if (id.substr(0,1)!="-")
				{
					showSelFolderItems(id);
		         }
			}
            }},
			{text: txtRefresh,
            listeners:{
            click: function(b,e){
            buildFolderList(true, tagsToo);
			var sel = Ext.getCmp("selFoldCombo");
			var ob = sel.getStore();
			ob.removeAll();
			var allObjs = new Array();
			if( tagsToo)allObjs = folderListObjs.concat(tagListObjs);
			else allObjs = folderListObjs.slice(0);
			ob.loadData(allObjs);
			sel.clearValue();
			}
            }},{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            listeners:{
            click: function(b,e){
            win.hide();
            win.destroy();}
            }}

        ]
    });

    showOLViewCtl(false);
	win.show();
	if(xdata != null)
	{
		var strg = "";
		if( xdata.length == 1){
			updateMailItemDisplay(xdata[0][1]);
			strg = mailTpl.apply(xdata[0][1].data);
		}else{
		for(var i=0; i < xdata.length; i++){
			strg += thistemp.apply(xdata[i][1].data);
		}
		}
		Ext.getCmp("selmtext").body.dom.innerHTML = strg;
		Ext.getCmp("selmtext").body.setStyle("font-size",jello.mailPreviewFontSize);

    }
	win.setActive(true);
  setTimeout(function(){Ext.getCmp("selFoldCombo").focus(true,100);},1);

	//shortcut keys
var map = new Ext.KeyMap('selectfold', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});


}
function selFolderAction()
{
var c = Ext.getCmp("selFoldCombo");
var id = c.folId;
if (id.substr(0,1)!="-")
{
var f = c.folFunc;
var d = c.folData;
var facts = Ext.getCmp("checkActs");
//var cked = null;
//var rked = null;
//var aked = null;
var xacts = {task:null,reply:null,appt:null, fwd: null, del: null, copy: null, cats: ""};
if( typeof(facts) != "undefined") {
	var factvals = facts.getValue();
	for(var i=0; i< factvals.length; i++)
	{
		var lbl = factvals[i].getId();
		switch(lbl){
		case "ckTask": xacts.task= factvals[i].getValue();break;
		case "ckRepl": xacts.reply = factvals[i].getValue();break;
		case "ckAppt": xacts.appt = factvals[i].getValue();break;
		case "ckFwd": xacts.fwd = factvals[i].getValue();break;
		case "ckDel": xacts.del = factvals[i].getValue();break;
		case "ckCopy": xacts.copy = factvals[i].getValue();break;
		}
	}
}
var mf = null;
if(( mf = Ext.getCmp('mtagfield')) != null && mf != "undefined") xacts.cats = mf.getValue();
if( xacts.task == true && xacts.cats ==""){
	Ext.Msg.alert("",txtCategoryNeeded);
	Ext.getCmp('mtags').render();
	Ext.getCmp('mtags').focus(null,100);
	return false;
}

if( d == null)
{	f(id);}
else
{	//f(id,d,cked, rked, aked);
	f(id,d,xacts);
}

} else{
 //open tag
 var tgid=parseInt(id);
 tgid=tgid+((-tgid)*2);

  showContext(tgid,1);

}
 setTimeout(function(){
		var w = Ext.getCmp("selectfold");
				w.destroy();
		},60);

}

function showSelFolderItems(id)
{
	var ibreader = new Ext.data.ArrayReader({}, [
		{name: 'subject'},
        {name: 'entryID'},
		{name: 'importance'},
		{name: 'sender'},
		{name: 'to'},
		{name: 'cc'},
		{name: 'body'},
		{name: 'attachment'},
		{name: 'attachmentList'},
		{name: 'created',type:'date'},
		{name: 'due',type:'date'},
		{name: 'iclass'},
		{name: 'unread',type:'boolean'},
		{name: 'icon'},
		{name: 'type'}
]);
	var goStore = new Ext.data.ArrayStore({
        id:'gotoStore',
		reader: ibreader,
		//reader: tabReader,
        fields: [
           {name: 'subject'},
           {name: 'entryID'},
		   {name: 'importance'},
		   {name: 'sender'},
		   {name: 'to'},
		   {name: 'cc'},
		   {name: 'body'},
		   {name: 'attachment'},
		   {name: 'attachmentList'},
		   {name: 'created',type:'date'},
		   {name: 'due',type:'date'},
		   {name: 'iclass'},
		   {name: 'unread',type:'boolean'},
		   {name: 'icon'},
		   {name:  'tags'},
		   //{name: 'groupon'}

        ],
		//groupField: "groupon",
		sortInfo:{field: 'created', direction: "DESC"}
    });
	var messageMark= [
           '<div onclick="openInboxItem(\'{entryID}\')">',
		   '<tpl if="typeof(values.to) != \'undefined\'">',
           'To: {to}\n<br>',
            '</tpl>',
            'From: {sender}\n<br>',
    		txtSubject+ ':{subject}\n<br>',
			'</div><hr>'
       ];
	var thistemp = new Ext.XTemplate(messageMark);

	try{
	var it = NSpace.GetFolderFromID(id);
	getInboxItems(1,goStore,it,"",0,true);
	var strg = "";
	var count = goStore.getCount();
	for(var i=0; i < count; i++){
			strg += thistemp.apply(goStore.getAt(i).data);
		}
	Ext.getCmp("selmtext").body.dom.innerHTML = strg;
	Ext.getCmp("selmtext").body.setStyle("font-size",jello.mailPreviewFontSize);
	}catch(e){}
}
function buildOLFolderList(topobj, list)
{
	try{
  var x = new Array(topobj.id, topobj.name.FolderPath.toLowerCase());
	//topobj.folderPath = topobj.name.FolderPath;
	list.push(x);
	if( topobj.name.Folders.Count != 0)
	{
		var child = enumerateFolders(topobj.id);
		for( var i=0; i < child.length; i++)
			buildOLFolderList(child[i],list);

	}
	}catch(e){}
}

function updateOLCategory(cpName,newName,del)
{
//Update outlook categories from jello contexts/projects (OL2007 only)

if (OLversion>11 && jello.updateOutlookCategories==true)
{
if (notEmpty(newName)==false){newName=null;}
  var cExists=false;var cPos=0;
  for (var x=1;x<NSpace.Categories.Count;x++)
  {
     if (trimLow(NSpace.Categories.Item(x).Name)==trimLow(cpName)){cExists=true;cPos=x;}
  }

  //category does not exist. Add it.
  if (cExists==false && newName==null){try{NSpace.Categories.Add(cpName);return true;}catch(e){return false;}}
  if (cExists==true && newName!=null){try{NSpace.Categories.Item(cPos).Name=newName;}catch(e){}}
  if (cExists==true && del==true){try{NSpace.Categories.Remove(cPos);}catch(e){}}

}
}
function updateInboxCount()
{
	var nInItems = NSpace.GetFolderFromID(jello.inboxFolder).Items.Count;
	inboxICount.innerHTML = "("+ nInItems.toString() +")";
}

function isOLReccuring(id)
{
//is outlook task recurring?
var t=NSpace.GetItemFromID(id);
return t.isRecurring;
}

