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

var thisDefinition=0;
var cDebug = false;
var impSetStore=new Ext.data.JsonStore({
        fields: [
           {name: 'id',type:'int'},
           {name: 'name'},
           {name: 'folderid'},
           {name: 'filter'},
           {name: 'removeoriginal',type:'boolean'},
           {name: 'createtags',type:'boolean'},
           {name: 'tagnew'},
           {name: 'noexistingcats',type:'boolean'},
           {name: 'ignoredone',type:'boolean'},
           {name: 'replacedups',type:'boolean'},
           {name: 'archive',type:'boolean'},
           {name: 'sfilter'},
           {name: 'cfilter'},
           {name: 'autostart',type:'boolean'},
           {name: 'alerts',type:'boolean'}
        ],
        data:jello.importSettings
    });

//define import settings record
 var impSetRecord = Ext.data.Record.create([
           {name: 'id',type:'int'},
           {name: 'name'},
           {name: 'folderid'},
           {name: 'filter'},
           {name: 'removeoriginal',type:'boolean'},
           {name: 'createtags',type:'boolean'},
           {name: 'tagnew'},
           {name: 'noexistingcats',type:'boolean'},
           {name: 'ignoredone',type:'boolean'},
           {name: 'replacedups',type:'boolean'},
           {name: 'archive',type:'boolean'},
           {name: 'sfilter'},
           {name: 'cfilter'},
           {name: 'autostart',type:'boolean'},
           {name: 'alerts',type:'boolean'}

           ]);


var iDasl="";

function pCollect()
{
thisGrid="collectArea";
isCollectScreen=true;
var ppnl=Ext.getCmp("portalpanel");
initScreen(true,"pCollect()");
//the collect process screen

var coarea = new Ext.form.TextArea({
 id: 'collectArea',
 height:panelHeight-jello.actionPreviewHeight-108,
 width:ppnl.getInnerWidth()-3,
// grow:true,
 emptyText:txtCollectHelp,
 style:{border:'none'},
 enableKeyEvents:true,
 listeners:{
		    keydown: function(f, e){
        if(e.getKey()==e.ENTER){if (!jello.collectAdvancedMode)validateLine();}
            }
		}
});

var ctbar=addCollectToolbar();

var sdd=getSavedDefs();

var copanel=new Ext.Panel({tbar:ctbar,items:[coarea]});

    ppnl.add(copanel);
    ppnl.setAutoScroll(false);
    ppnl.doLayout();
    var prep=Ext.getCmp("previewpanel");
    prep.update(sdd);
{ppnl.setTitle("<img src=img\\collect16.png style=float:left;> "+txtCollect);}

var cmap0 = new Ext.KeyMap(document, {
    key: "t",
    ctrl:true,
    fn: function(){
	Ext.getCmp('cotags').focus();
            },
    scope: this});

var cmap1 = new Ext.KeyMap(document, {
    key: "d",
    ctrl:true,
    fn: function(){
	Ext.getCmp('popdate').showMenu();
            },
    scope: this});

//ctrl+o performs collect
var cmap2 = new Ext.KeyMap(document, {
    key: "o",
    ctrl:true,
    stopEvent:true,
    fn: function(){
	doTextCollect();
    },
    scope: this});

Ext.getCmp("collectArea").focus(false,true);
resizeGrids(Ext.getCmp("collectArea"),100);
}


function insertTag()
{
var tg=Ext.getCmp("cotags").getValue();
var tar=Ext.getCmp("collectArea");
var tt=tar.getValue();
if (!jello.collectAdvancedMode){if (tt.substr((tt.length-1),1)!=","){tt+=",";}}
else{tt+=" tags= "; if( tg.indexOf(' ') != -1) tg = '"' + tg + '"';}

tar.setValue(tt+tg);
var ntar=tar.getValue();
  setTimeout(function(){tar.focus();tar.selectText(ntar.length);},3);
}

function addCollectToolbar()
{
getTagsArray();

    var dateMenu = new Ext.menu.DateMenu({
        handler : function(dp, date){
            var tt=Ext.getCmp("collectArea").getValue();
        			if(jello.collectAdvancedMode)
        				tt += " ";
        			else
                if (tt.substr((tt.length-1),1)!=","){tt+=",";}
              var dt=date.format('m/j/Y');
              if(jello.dateFormat == 0){dt=date.format('j/m/Y');}
              if(jello.dateFormat == 2){dt=date.format('Y/m/d');}
              Ext.getCmp("collectArea").setValue(tt+dt);
              Ext.getCmp("collectArea").focus();
        }
    });

    cbar = new Ext.Toolbar();

        cbar.add(new Ext.form.ComboBox({
                fieldLabel: txtTags,
                id: 'cotags',
                store:globalTags,
                hideTrigger:false,
                typeAhead:true,
                triggerAction:'all',
                //editable:false,
                disableKeyFilter:true,
                emptyText:txtInsTag,
                mode:'local',
                listeners:{
            select: function(cb,rec,idx){
            insertTag();
            }
		}

            }),{
            icon: 'img\\calendar.gif',
            cls:'x-btn-icon',
            id:'popdate',
            tooltip: txtInsDate+' [Ctrl+D]',
            menu: dateMenu
            },'-',
        {
        icon: 'img\\folder.gif',
        cls:'x-btn-icon',
        tooltip: txtImpOLFolder,
			listeners:{
            click: function(b,e){
			e.stopEvent();
			e.preventDefault();
			e.stopPropagation();
			collectOLFolder();
            }
            }
        },'-',{
        text:"<font size=2>"+txtBtnColText+"</font>",
        tooltip: txtCollectBtnInfo+' (Ctrl+O)',
        iconCls:'colectbuticon',
	    id:'ccol',
		listeners:{
            click: function(b,e){
			doTextCollect();
            }
            }
             },'-',

             {
        text:"<font size=2>"+txtInbox+"&nbsp;("+countInboxItems()+")</font>",
        tooltip: txtProcInbox,
	    id:'cinb',
	    tooltip:txtProcInboxInfo,
		listeners:{
            click: function(b,e){pInbox();}
            }
        },
             {
		icon: "img\\icon_link.gif",
		tooltip: "Make a multi-line entry",
	    id:'cjoin',
	    //disabled: (jello.collectAdvancedMode? false:true),
		hidden: (jello.collectAdvancedMode? false:true),
		listeners:{
            click: function(b,e){
            doJoinSelected();
            }
            }
        },'->','-',{
        text: txtColAdvanced,
        id:'vadv',
        enableToggle:true,
        pressed:jello.collectAdvancedMode,
        toggleGroup:'ccol',
        handler: toggleAdvancedMode
        },'-',{
        icon: 'img\\icon_info.gif',
        cls:'x-btn-icon',
        tooltip: '<b>'+txtHelp+'</b>',
			listeners:{
            click: function(b,e){
			e.stopEvent();
			e.preventDefault();
			e.stopPropagation();
			collectHelp();
            }
            }
        }

         );
return cbar;
}

function collectOLFolder(pfid,pfd,parc,prem,pcre,pcim,pdim,pdon,pdup,pflt,psflt,pcflt)
{
//outlook folder items import/collect form
thisDefinition=0;
if (notEmpty(pfid)){var folderID=pfid;}else{var folderID=null;}
try{var folderName=NSpace.GetFolderFromID(pfid).FolderPath;}catch(e){var folderName="";}
if (notEmpty(parc)==false){parc=false;}
if (notEmpty(prem)==false){prem=false;}
if (notEmpty(pcre)==false){pcre=true;}
if (notEmpty(pdim)==false){pdim=false;}
if (notEmpty(pdon)==false){pdon=false;}
if (notEmpty(pdup)==false){pdup=false;}

//select an outlook folder to collect items
var simple = new Ext.FormPanel({
        labelWidth: 75, // label settings here cascade unless overridden
        frame:true,
        iconCls:'olcollectformicon',
        title: txtImpOlFldrInfo,
        width:400,
        bodyStyle:'padding:0px 0px 0 0px',
        floating:false,
        id:'olcolform',
        buttonAlign:'center',
        defaults: {width: 270},
        defaultType: 'textfield',
        items: [
        	new Ext.form.TriggerField({
	fieldLabel:txtOLFolder,
	value:folderName,
	onTriggerClick:function(e){selectColFolder(this);},
	labelStyle:'width:150px',
	id:'folsel'
	}),

				new Ext.form.Hidden({
				id:'folid',
				value:folderID
				}),

	{
                fieldLabel: txtSubjFltr,
                id: 'sfil',
                allowBlank:true,
                value:psflt,
                labelStyle:'width:150px',
                listeners:{blur: function(b,e){
					updateColStatus();}
        }
            },
        {
        fieldLabel: txtCatFltr,
        id: 'cfil',
        emptyText:txtCatFltrInfo,
        labelStyle:'width:150px',
        allowBlank:true,
        value:pcflt,
        listeners:{blur: function(b,e){
					updateColStatus();}
        }
    },
    new Ext.form.Checkbox({
	fieldLabel:txtColIgnDone,
	checked:pdon,
	labelStyle:'width:150px',
	hidden:true,
	hideLabel:false,
        listeners:{check: function(b,e){
					updateColStatus();}
		},
	id:'mdon'
	}),
    new Ext.form.Label({
				html:'',
				height:40,
				id:'folstatus'
				}),
           {
        fieldLabel: txtColTagImpItms,
        id: 'cas',
        emptyText:txtTagFltrInfo,
        labelStyle:'width:150px',
        value:pcim,
        allowBlank:true

    },
        new Ext.form.Checkbox({
	fieldLabel:txtColTagCreItms,
	checked:pcre,
	labelStyle:'width:150px',
	width:25,
	hidden:false,
	hideLabel:false,
	id:'mcre'
	}),
    new Ext.form.Checkbox({
	fieldLabel:txtColArcCreItms,
	checked:parc,
	labelStyle:'width:150px',
	width:25,
	hidden:true,
	hideLabel:false,
	id:'marc'
	}),

    new Ext.form.Checkbox({
	fieldLabel:txtColRmvTskItms,
	checked:prem,
	labelStyle:'width:150px',
	hidden:true,
	hideLabel:false,
	id:'mrem'
	}),
    new Ext.form.Checkbox({
	fieldLabel:txtColIgnCatsItms,
	checked:pdim,
	labelStyle:'width:150px',
	hidden:false,
	hideLabel:false,
	id:'mdim'
	}),    new Ext.form.Checkbox({
	fieldLabel:txtColNoDupItms,
	checked:pdup,
	labelStyle:'width:150px',
	hidden:false,
	hideLabel:false,
	id:'mdup'
	})


        ]
    });

   var win = new Ext.Window({
        title: txtCollect,
        width: 480,
        height:400,
        id:'olcollectform',
        minWidth: 300,
        minHeight: 200,
        resizable:true,
        defaultButton:'sbut',
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: simple,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        modal:true,
        buttons: [{
            text:"<b>"+txtCollect+"</b>",
            id:'sbut',
            tooltip:txtCollect+' [F2]',
            listeners:{
            click: function(b,e){
            var fid=Ext.getCmp("folid").getValue();
            var arc=Ext.getCmp("marc").getValue();//archive
			var rem=Ext.getCmp("mrem").getValue();//remove original task
			var cre=Ext.getCmp("mcre").getValue();//create non existing tags
			var cim=Ext.getCmp("cas").getValue();//tag newly created actions with this list of tags
			var dim=Ext.getCmp("mdim").getValue();//do not import existing categories of items
			var don=Ext.getCmp("mdon").getValue();//ignore done tasks
			var dup=Ext.getCmp("mdup").getValue();//replace duplicates


            importOLFolder(fid,iDasl,arc,rem,cre,cim,dim,don,dup);
            }
            }
        },{
            text: txtColSaveDef,
            tooltip:txtColSaveDefInfo,
                        listeners:{
            click: function(b,e){
            saveColDef();
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



//shortcut keys
var map = new Ext.KeyMap(document, {
    key: Ext.EventObject.F12,
    fn: function(){

            try{win.destroy();}catch(e){}

            },
    scope: this});

var map2 = new Ext.KeyMap(document, {
    key: Ext.EventObject.F2,
    fn: function(){try{}catch(e){}},
    scope: this});

//show the form
    showOLViewCtl(false);
    win.show();
    win.setActive(true);

}

function selectColFolder(el)
{
//popup outlook folder selection to collect from
var ss="";
var t=NSpace.PickFolder();
	if (t!=null)
	{
   if (t.EntryID==jello.inboxFolder|| t.EntryID==jello.archiveFolder || t.EntryID==jello.actionFolder){alert(txtMsgNoSystemFolder);return;}

  		el.setValue(t.FolderPath);
  		Ext.getCmp("folid").setValue(t.EntryID);
  		updateColStatus(t);
	}

}

function updateColStatus(t)
{
//update counter of items to be imported
if (t==null)
{
  var id=Ext.getCmp("folid").getValue();
  if (notEmpty(id)==false){return;}
  var t=NSpace.GetFolderFromID(id);
}
var simple=Ext.getCmp("olcolform");
Ext.getCmp("marc").hide();Ext.getCmp("marc").setValue(false);
Ext.getCmp("mrem").hide();;Ext.getCmp("mrem").setValue(false);
Ext.getCmp("mdon").hide();

if (t.DefaultItemType==0)
{//mail folder selected
Ext.getCmp("marc").show();
}

if (t.DefaultItemType==3)
{//task folder selected
Ext.getCmp("mrem").show();
Ext.getCmp("mdon").show();
}

var cf=Ext.getCmp("cfil").getValue();
var sf=Ext.getCmp("sfil").getValue();
var don=Ext.getCmp("mdon").getValue();


iDasl="";var tcc="";

if (notEmpty(sf))
{
iDasl="(urn:schemas:httpmail:subject LIKE '%"+sf+"%')";
}

if (notEmpty(cf))
{
//test filter
var cc=cf.split(",");
if (notEmpty(iDasl)){iDasl+=" OR ";}
iDasl+="(";

	for (var x=0;x<cc.length;x++)
	{
		if (notEmpty(Trim(cc[x])))
		{
		iDasl+="(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + Trim(cc[x]) + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + Trim(cc[x]) + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + Trim(cc[x])+ "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + Trim(cc[x]) + "') ";
		if (x<(cc.length)-1){iDasl+="OR ";}else{iDasl+=")";}
		}
	}

}

if (don==true && t.DefaultItemType==3)
{if (notEmpty(iDasl)){iDasl+=" AND ";}
iDasl+="http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";}

if (notEmpty(iDasl))
{
var tc=t.Items.Restrict("@SQL="+iDasl);
if (tc.Count>0){tcc="("+txtFilterTo+" "+tc.Count+")";}else{tcc="("+txtColAlFltNoItms+")";}
}

folstatus.innerHTML="<p align=right>"+t.Items.Count()+" Items "+tcc+"</p><p align=right></p>";
Ext.get("folstatus").fadeIn();
}

function importOLFolder(fid,dasl,arc,rem,cre,cim,dim,don,dup)
{
//import items from a folder

//arc=archive item after creation
//rem=remove original task (for task folders only)
//cre=create not existing tags
//cim=add those tags to newly created items
//dim=don't import existing categories of original item
//don=ignore done tasks (for task items)
//dup=replace duplicates

if (notEmpty(fid)==false){alert(txtMsgNoFolder);return;}

	if(arc==true)
	{
	var arcfolder=getDefaultArcFolder();

   if (notEmpty(arcfolder)==false){alert(txtNoArchiveFolder);return;}
   else
   {var arcTo=NSpace.GetFolderFromID(arcfolder);}
	}



//GET OUTLOOK FOLDER
	 var t=NSpace.GetFolderFromID(fid);
	 var ttype=t.DefaultItemType;
	 var taskfol=NSpace.GetFolderFromID(jello.actionFolder);
	 var newoi;
	 var icount=0;
	 try{var addTags=Ext.getCmp("cas").getValue();}catch(e){var addTags="";}

//apply filters
	 if (notEmpty(dasl)){var tt=t.Items.Restrict("@SQL="+dasl);}else{var tt=t.Items;}

	          var preItms=getPreviewTitlesFromDef(tt);
            var cmsg=txtColImpResult;
            cmsg=cmsg.replace("%1",tt.Count).replace("%2",preItms);
            var choice=confirm(cmsg);
            if (choice==false){return;}

	if (tt.Count>0)
	{
		for (var x=tt.Count;x>=1;x--)
		{
			var oi=tt.Item(x);
			var origDate=new Date(oi.LastModificationTime);
			var origParent=oi.Parent;
			var newoi=oi;
			status=txtStatusGetting+" "+oi;
      icount++;
      var cats="";
      if (dim==false || dim=="false"){
      try{
      cats=oi.itemProperties.item(catProperty).Value+";";
      }catch(e){alert(txtAlNoDwnldFol);return;}
      }else{var cats="";}

			if (ttype==3)
			{
			//task items are just moved or copied to the destination folder
				if (rem==true || rem=="true")
				//remove original task
				{newoi=oi.Move(taskfol);}
				else
				//copy original task
				{var newoione=oi.Copy();newoi=newoione.Move(taskfol);}
				var newcats=cats+";"+cim;
				newoi.itemProperties.item(catProperty).Value=newcats;
				newoi.Save();
			//check for categories if not exists as tags
			if (cre==true){addTags+=getUniqueImportedCategories(newcats,addTags);}
			if (dup==true){checkDuplicates(newoi,origDate,origParent);}

      }

			else
			{
			//non task items
				if (arc==true)
				{
				//archive item
				newoi=oi.Move(arcTo);
				}

			//create task out of item
			var newA=createActionOL(newoi.Subject);
			newA.itemProperties.item(catProperty).Value=cats+cim;
			//check for categories if not exists as tags
			if (cre==true && dim==false){addTags+=getUniqueImportedCategories(cats,addTags);}
			//attach original item to task
			setAttachmentProperty(newA,newoi.EntryID);
			newA.Save();

			if (dup==true){checkDuplicates(newA,origDate,origParent);}
			}

		}
	}

	var ntcount=0;
	status="";
	if (cre==true){ntcount=checkNewTags(addTags);}
	var cmsg=txtAlColSucc;
	cmsg=cmsg.replace("%1",icount).replace("%2",ntcount).replace("%3","<a class=jellolinktop onclick=pInbox();><b>Inbox</b></a>");
	Ext.info.msg(txtCompleted,cmsg);
	Ext.get("topInbox").highlight();
	try{Ext.getCmp("olcollectform").destroy();}catch(e){}
	if (icount>0){Ext.getCmp("cinb").setText("<font size=2 color=blue>"+txtInbox+"&nbsp;("+countInboxItems()+")");}
}

function getUniqueImportedCategories(cats,addTags)
{
//get imported outlook item categories.
//If they are not existing as tags in jello system add them to the "tags to be created" array
//by returning the new unique tags divided by ;
if (notEmpty(cats)==false){return "";}
cats=cats.replace(new RegExp(",","g"),";");
var ret="";
try{var c=cats.split(";");}catch(e){return "";}

var t=addTags.split(",");
	for (ct in c)
	{
	var addIt=true;
			for (tt in t)
			{
				if (trimLow(t[tt])==trimLow(c[ct])){addIt=false;}
			}
	if (addIt==true){ret+=c[ct]+",";}
	}

	return ret;
}

function setAttachmentProperty(it,val,append){
//set custom attachment field value.
var jn="";
try{
jn=it.UserProperties.Item("OLID").Value;
}catch(e){it.UserProperties.Add("OLID",1,true);}
  if (append && notEmpty(jn)){val=jn+";"+val;}
  it.UserProperties.Item("OLID").Value=val;
}
function validateLine()
{
//validate line entered for collection
var tt=Ext.getCmp("collectArea").getValue();
var t=tt.split("\n");
var lline=lastNonBlankLine(t);
	if (notEmpty(t[lline])==false){return;}

	//get last line and split to tags
	var aLine=t[lline];
    var o=aLine.charCodeAt(aLine.length-1);
    if (o==13){aLine=aLine.substr(0,aLine.length-1);}
	var line=aLine.split(",");

		if (line.length>1)
		{
			var ret=line[0];
		 for (var x=1;x<line.length;x++)
		 {
			if (notEmpty(line[x])){ret+=","+validateTag(line[x]);}
		 }
			t[lline]=ret;
			var nret="";
			for (var x=0;x<t.length;x++)
			{
       nret+=t[x];
			 nret+="\n";
			}
			nret=nret.substr(0,nret.length-1);
			Ext.getCmp("collectArea").setValue(nret);
		}
}

function validateTag(tg)
{
//search for tags starting with criteria

//if its a date leave it alone
//tagStore.filter("tag",tg);
tagStore.filterBy(function ff(r){
var a=r.get("istag");
var b=r.get("tag");
var c=trimLow(b.substr(0,tg.length))==trimLow(tg);
if (a==true && c==true){return true;}else{return false;}
});
if (tagStore.getCount()>0){return tagStore.getAt(0).get("tag");}
tagStore.clearFilter();
return tg;
}

//*** new code

// check to see if a string is a collect Date
// collect dates can have the form date+[starttime[+endtime]]
// retun an Array
// 0 = is it a date, 1 = date, 2 = starttime, 3 = endtime
// 2 and 3 null if not present
function collectApptDate(str)
{

	var vals = new Array();
	// if no + than just a single date
	// or if plus at the start
	vals[3] = vals[2] = null;

	var plus;
	if( (plus=str.indexOf('+')) <= 0)
	{
		vals[0] = isDate(str);
		if( vals[0]){
			vals[1] = str;
			return vals;
		}
	}

	// has a plus, now see if it is a Date
	var xdate = str.substring(0,plus);
	if( (vals[0] = isDate(xdate)) == false)
		return vals;
	// ok we have a Date
	vals[1] = xdate;
	// get part after the +, if nothing return what we have
	str = str.substring(plus+1);
	if(str.length == 0 || str =="")
		return vals;
	// if plus there is also an end time
	if((plus = str.indexOf('+')) <0)
		vals[2] = str;
	else{
		vals[2] = str.substring(0,plus);
		vals[3] = str.substring(plus+1);
	}
	return vals;

}
var CollectItemTypes = new Array(
	["appt", "a"],
	["apt", "a"],
	["meet", "a"],
	["cal", "a"],
	["task", "t"],
	["todo", "t"],
	["tsk", "t"],
	["do", "t"],
	["action", "t"],
	["act", "t"],
	["note", "n"],
	["notes", "n"],
	["memo", "n"],
	["jot", "n"],
	["person","p"],
	["contact","p"],
	["pers", "p"]

);

// does a string contain one of the elements of an Array
// the argument col tells us which element of the Array
// we should try to match.  pos tells us if the match must occur at
// a specific point in the string.  0 is the start , 1, second char
// -1 means string must be at the end and null means no requirement
// returns an array , elem 0 = aray element index of match,
// elem 1 = position of the match
// NOTE:  List must be ordered so that if one entry is a substring
// of another entry, the shorter entry is first.  So cat  must
// precede category

function isElementInString( str, arr, col, pos)
{
	var p;
	var x = new Array();
	x[0] = -1;
	var multi=false;
	if( typeof(arr[0]) == "object")
		multi = true;
	for(var i=0; i < arr.length; i++)
	{
		var e;
		if(multi)
			e = arr[i][col];
		else
			e = arr[i];

		if( (p = str.indexOf(e) ) != -1)
		{
			if( pos == null || (pos != -1 && p == pos )
				|| (pos == -1 && (p + e.length)  == str.length)
			)
			{
				x[0] = i; x[1] = p;
			}
		}

	}
	return x;

}
function getCollectType(str)
{
	var vals = new Array();
	// set defaults
	// task and just the string we are passed
	vals[0] = "t";	//task
	vals[1] = str;
	// if no : ort : at 0, then not an item type specified
	var colon = -1;
	if( (colon = str.indexOf(':') ) <= 0)
		return vals;
	// get the possible type and separate from rest of string
	var type = str.substring(0,colon);

	for( var i =0 ; i < CollectItemTypes.length; i++)
	{
		if( CollectItemTypes[i][0] == type)
		{
			vals[0] = CollectItemTypes[i][1];
			vals[1] = "";
			break;
		}
	}
	return vals;

}
function LineObject()
{
	this.type = "t";
	this.desc = "";
	this.startDate = this.endDate = "";
	this.startTime = "";
	this.endTime = "";
	this.allDay = false;
	this.email = new Array();
	this.phone = new Array();
	this.company="";
	this.title="";
	this.address = "";
	this.homeAddress = "";
	this.body = "";
	this.cats = "";
	this.state = "";
	this.lastWord = "";
	this.lastWordType = "";
	this.quoteBag = "";


}
LineObject.prototype.toString = function(user){
	var u = user;
	if (u == null)
		u = false;
	var ret = "";
	if( !u) ret += this.type + "\n";
	ret += this.desc + "\n";
	ret += "start: "+ this.startDate + " - " +this.startTime + "\n";
	ret += "end: " + this.endDate + " - " +this.endTime + "\n";
	ret += "allday is " + this.allDay + "\n";
	ret += "company is " + this.company + "\n";
	ret += "title is " + this.title + "\n";
	ret += "address is " + this.address + "\n";
	ret += "home address is " + this.homeAddress + "\n";
	for (var i=0; i < this.email.length; i++)
		ret += "email: " + this.email[i][0] + " : " + this.email[i][1] + "\n";
	for (var i=0; i < this.phone.length; i++)
		ret += "phone: " + this.phone[i][0] + " : " + this.phone[i][1] + "\n";
	ret += "Body : " + this.body + "\n";
	ret += "categories: "+ this.cats +  "\n";
	if( !u) ret += "state: "+ this.state +  "\n";
	return ret;
};

LineObject.prototype.checkType = function(arg){
	var x = new Array();
	x = getCollectType(arg);
	this.type = x[0];
	if( x[1] != null && x[1] != "")
		this.desc += x[1];
	return;
};
LineObject.prototype.noiseToDesc = function()
{
	if( this.lastWord == "" && this.quoteBag == "")
		return;
	if( this.lastWordType == "datenoise"){
		var dx = new Date();
		var res = datenoiseToDate(dx,this.lastWord);
		if( res != null){
			this.setDate(res);
			this.lastWord = "";
		}
	}
	if( this.lastWord == "")
		this.lastWord = this.quoteBag;
	else if( this.quoteBag != "")
		this.lastWord += " " + this.quoteBag;
	if (this.desc != "")
		this.desc += " ";
	this.desc += this.lastWord;
	this.clearLast();
};

LineObject.prototype.clearLast = function()
{
	this.lastWord = this.lastWordType  = this.quoteBag = "";
};

var txtCategoryTags="tags,cat,category,categories";
var txtNoteTags="note,notes,body";

LineObject.prototype.nextWord = function(arg, newTags)
{
	if( cDebug)
		var debugThis = this;
	var nd;
	arg = Trim(arg);
	if( this.state.match(/q/i) == null)
	{

		var arlist = txtCategoryTags.split(",");
		var aret = isElementInString( arg,arlist,0,0);
		if( aret[0] != -1 && aret[1] == 0 && arg.charAt(arlist[aret[0]].length) == "="){
		// categories to end of line

			if( this.state.indexOf("N") >=0 )
				this.state = this.state.replace(/N/g,"");
			this.state += "T";
			this.noiseToDesc();
			aret = arg.split("=");
			if( aret.length != 1 && aret[1] != "")
				arg = Trim(aret[1]);
			else
				return;
		}

		arlist = txtNoteTags.split(",");
		aret = isElementInString( arg,arlist,0,0);
		if( aret[0] != -1 && aret[1] == 0 && arg.charAt(arlist[aret[0]].length) == "="){
			// categories to end of line
			if( this.state.indexOf("T") >= 0)
				this.state = this.state.replace(/T/g,"");
			this.state += "N";
			this.noiseToDesc();
			aret = arg.split("=");
			if( aret.length != 1 && aret[1] != "")
				arg = Trim(aret[1]);
			else
				return;
		}


		if( !(this.state.indexOf("N") < 0)){
			if( this.body != "")
				this.body += " ";
			this.body += arg;
			return;
		}
	}
	// don't allow quotes in notes, all passes through
	if( this.type != "n"){
		var quot = arg.indexOf('"');
		if( (quot == 0 && arg.length != 1	) || (quot == 0 && arg.length == 1 && this.state.indexOf("Q") < 0 )){ // initial quote means turn on
			//this.noiseToDesc();
			this.state += "Q";
			arg = arg.substring(1);
		}else if( quot == arg.length -1 ){
			this.state = this.state.replace("Q","q");
			arg = arg.substring(0,arg.length-1);
		}
		if( this.state.match(/q/i)){
			// in quoting mode, add the arg

			if( arg == "")
				return;
			if( this.quoteBag != "")
				this.quoteBag += " ";
			this.quoteBag += arg;
			arg ="";
			// if lower q then we had a closing quote, reset state
			if( this.state.indexOf("q") >= 0 ){
				this.state = this.state.replace(/q/i,"");
			}else
				return;
		}
	}
	// are we processing tags?
	if( !(this.state.indexOf("T") < 0)){
		// this is a tag
		var xarg = arg;
		if( this.quoteBag != "")
			xarg = this.quoteBag;
		var r=getTag(xarg);
		tagStore.filter("tag",xarg);
		//if (tagStore.getCount()==0 && isDate(el)==false)
		var dt = collectApptDate(xarg);
		if( dt[0] == false){
			if (tagStore.getCount()==0)
				if (!arrayExists(xarg,newTags)){newTags.push(xarg);}
			if(this.cats != "")
				this.cats += ";";
			this.cats += xarg;
			tagStore.clearFilter();
		}
		this.quoteBag = "";
		return;
	}
	time_re = /^\d{1,2}:\d{2}([AapP][mM])?$/;
	// if a note, everything except tags goes in the note
	//if( this.type == "n"){
	//	if( this.desc != "")
	//		this.desc += " ";
	//	this.desc += arg;
	//	return;
	//}else
	if(this.type == "a" && arg.match(time_re)  ) // might be a time
	{

		if( this.lastWordType != "timenoise" && this.lastWordType != "separator")
			this.noiseToDesc();

		this.clearLast();
		if( this.startTime != "" && this.endTime != "")
		{
			if( this.desc != "")
				this.desc += " ";
			this.desc += this.startTime;
			this.startTime = this.endTime;
		}
		if(this.startTime == "" ){
			this.startTime = arg;
			this.state +="S";
		}else if( this.startTime != ""  ){
			this.endTime = arg;
			this.state +="E";
		}
		return;
	}else if( (this.type == "a" || this.type == "t") &&
	( nd=(cIsDate(arg, this.lastWordType =="datenoise"?this.lastWord:""))) != null)  // is it a date
	{
		var xd;
		if( this.lastWordType == "datenoise")
		{
			var sfmt = jelloDateFormatString();
			var xfmt = sfmt.substring(0,sfmt.length-1) + 'Y';
			var dx = Date.parseDate(nd,sfmt);
			xd = datenoiseToDate(dx,this.lastWord);
		}else
			this.noiseToDesc();
		this.setDate(xd==null?nd:xd);
		return;
	}else if(this.type == "p" && (isEmail(arg) || isPhoneNumber(arg))){
		if( this.lastWordType != "perstype")
			this.noiseToDesc();
		var xt = txtWork.toLowerCase();
		if( this.lastWordType == "perstype" && this.lastWord != "")
			xt = this.lastWord;
		var x = new Array();
		x[0] = xt;
		x[1] = arg;
		if( isEmail(arg))
			this.email.push(x);
		else
			this.phone.push(x);
		this.clearLast();
	} else if( this.type == "p" &&  this.lastWordType == "perstype"){
		var xarg = arg;
		if( this.quoteBag != ""){
			xarg = this.quoteBag;
			this.quoteBag = "";
		}
		if( this.lastWord == txtCompany.toLowerCase())
			this.company = xarg;
		else if ( this.lastWord == txtTitle.toLowerCase())
			this.title = xarg;
		else if ( this.lastWord == txtAddress.toLowerCase())
			this.address = xarg;
		else if ( this.lastWord == txtHomeAddress.toLowerCase())
			this.homeAddress = xarg;
		else if (xarg != ""){
			if( this.body != "")
				this.body += " ";
			this.body += xarg;
		}
		this.clearLast();
		return;
	}else if( this.type == "a" && arg == "allday" ){
		// allday after either start or end date
		this.noiseToDesc();
		this.allDay = true;
		this.clearLast();
	}else if( this.type == "p" &&
		(arg == txtHome.toLowerCase() || arg == txtWork.toLowerCase() || arg == txtMobile.toLowerCase() || arg == txtFax.toLowerCase() || arg == txtCompany.toLowerCase()  ||
		arg == txtTitle.toLowerCase() || arg == txtAddress.toLowerCase() || arg == txtHomeAddress.toLowerCase())){
		this.noiseToDesc();
		this.lastWord = arg;
		this.lastWordType = "perstype";
	}else if( arg == "-"){
		this.noiseToDesc();
		this.lastWord = arg;
		this.lastWordType = "separator";
	}else if ( (this.type == "a" || this.type == "t") && isDateNoise(arg)){
		if( this.lastWordType != "datenoise")
			this.noiseToDesc();
		if( this.lastWord != "")
			this.lastWord += " ";
		this.lastWord +=  arg;
		this.lastWordType = "datenoise";
	}else{
		// just a word
		this.noiseToDesc();
		if (this.desc != "")
		this.desc += " ";
		this.desc += arg;
		this.clearLast();
	}

};


LineObject.prototype.setDate = function(nd){
		nd = Trim(nd);
		var ndsp = nd.indexOf(" ");
		if( ndsp > 0)
			nd = nd.substring(ndsp+1);
		if( this.lastWordType != "datenoise" && this.lastWordType != "separator")
			this.noiseToDesc();
		var was="";
		// do we have a start date
		if(this.startDate != "" && this.endDate == ""){
			this.endDate = nd;
			this.state += "e";
		}else if(this.startDate == "" ){
			this.startDate = nd;
			this.state += "s";
		}else{
			// more than 2 dats, so first goes to description
			if( this.desc != "")
				this.desc += " ";
			this.desc += this.startDate;
			this.startDate = this.endDate;
			this.endDate = nd;
		}
		this.clearLast();
};

function cIsDate(str, datenoise)
{
	if( str == "")
		return null;
	// ie7 and 8 treat a single or 2 digit year as being in 1900's
	// fix this up
	// 	if( str.indexOf('/') != str.lastIndexOf('/'))
	// 	{
	// 		// ok, more than 1 separator
	// 		// grab last part
	// 		var dt = str.substring(str.lastIndexOf('/')+1);
	// 		var z = str.replace(/\//g,"");
	// 		// since the year might start with 0, need to say
	// 		// use base 10
	// 		var dy = parseInt(dt,10);
	// 		dx = /\D/
	// 		// if string is all numbers and year part is 1 or 2 digits
	// 		// fix it ot be 2000
	// 		if( !dx.test(z) && dt.length > 0 && dt.length < 3 )
	// 		{
	// 			str = str.substring(0,str.lastIndexOf('/')) + "/" + (dy+2000).toString();
	// 		}
	// 	}
	// check for numeruc formatted string
	// using specified format
	var sfmt = jelloDateFormatString();
  var xfmt = sfmt.substring(0,sfmt.length-1) + 'Y';
	var dx = Date.parseDate(str,sfmt);
	if( typeof(dx) != "undefined" && dx != null && dx != "")
  return dx.format(xfmt);
	// might be of the form with just month and day
	var dx = Date.parseDate(str,sfmt.substring(0,sfmt.lastIndexOf(jello.dateSeperator)));
	if( typeof(dx) != "undefined" && dx != null && dx != "")
	return dx.format(xfmt);
	var dx = Date.parseDate(str,xfmt);
	if( typeof(dx) != "undefined" && dx != null && dx != "")
		return dx.format(xfmt);
	// might be a 4 digit year


// 	if( isDate(str) )
// 		return str;
	// might be 2 digits, so shorten the format


// 	if( str.indexOf('/') >= 0)
// 	{
// 		var z = str.replace(/\//g,"");
// 		dx = /\D/
// 		if( dx.test(z))
// 			return null;
// 		z = str.split('/');
// 		if( z.length != 2 || z[0].length >2 || z[0].length == 0 ||
// 			z[1].length >2 || z[1].length == 0 )
// 			return null;
// 		var x = new Date();
// 		var y = str + "/" + x.getYear();
// 		if( isDate(y)){
// 			// date - make sure the year is right
// 			// we have month and day, if month is less than
// 			// current we need next year or if month is same and
// 			// date is less
// 			var xx = new Date(y);
// 			if( xx.getMonth() < x.getMonth() ||
// 				( xx.getMonth() == x.getMonth() && xx.getDate() < x.getDate()))
// 			{
// 				// month is befroe now, so date must be next year
// 				xx.setFullYear(xx.getFullYear() + 1);
// 				return xx.format(jelloDateFormatString());//xx.toDateString();
// 			}else
// 				return y;
//
// 		}else
// 			return null;
// 	}else{
		var slower = str.toLowerCase();
		var dx = new Date();
				if ( slower == txtToday.toLowerCase() ||
				slower == txtNow.toLowerCase())
			return dx.format(jelloDateFormatString());//dx.toDateString();
		if ( slower == txtTomorrow.toLowerCase()){
			dx.setDate(dx.getDate() + 1);
      return dx.format(xfmt);//dx.toDateString();
		}
				// if last word was next see if this is week, month, year
		var lw = Trim(datenoise);
		if( datenoise.toLowerCase() == txtNext.toLowerCase())
		{
			var dxn = null;
			// make next week mean
			if( slower == txtWeek.toLowerCase()){
				dxn = dx; dxn.setDate(dxn.getDate() + 7);
			}else if( slower == txtMonth.toLowerCase()){
				dxn = dx; dxn.setDate(dxn.getMonth() + 1);
			}else if( slower == txtYear.toLowerCase()){
				dxn = dx; dxn.setDate(dxn.getFullYear() + 1);
			}
			if( dxn != null)
				return dxn.format(xfmt);//dxn.toDateString();
		}


		// ok, now try for named dates
		try{
			var a=txtDayList.split(",");
			var slen = slower.length;
			for(var i=0; i<a.length; i++)
			{
				var alower = a[i].toLowerCase();
				if( slower != alower &&
					slower != alower.substring(0,slen))
					continue;
				var doffset = 0;

				var tday = dx.getDay();
				var nwd = Trim(datenoise);
				var narr = datenoise.split(" ");
				if( narr.length > 0)
					ndw = narr[length -1];
        nwd = nwd.toLowerCase();

					if( i > tday)
						doffset = i - tday;
					else if( i < tday)
						doffset = 7 - tday + i;
					// hard to figure out what folks mean by next
					// if week day and day is in week, then the following week
					// for example on tues next wed means the following week
					// on fri next mon means the one coming up, but not to all
					// on weekend next mon or tues... mean not the next few days, but
					// the following week
				if( nwd.match(txtNext.toLowerCase())){

				if( i >= tday )

						doffset += 7;
					}
					dx.setDate(dx.getDate() + doffset);

				return dx.format(xfmt);//dx.toDateString();

			}

		}catch(e){return null;}
// 	}
	return null;
}
// convert the date stuff to a Date
// might be in the form 1 week from tuesday, in 2 weeks, in 2 weeks from tues
// dx is the base date, datenoise is the text
// if force we found a real date word and need to add the noise
// to the date
function datenoiseToDate(dx,datenoise)
{
	var theCount = null;
	var dn = datenoise.toLowerCase();
	// get rid of in, from and next
	// next should have been handled when we matched the day of week
	var rex = new RegExp("(" + txtIn + "|" + txtFrom  + "|" + txtNext + ")","ig" );
	dn = dn.replace(rex,"");
	dn = dn.replace(/  */," ");
	dn = Trim(dn);
	var a = dn.split(" ");

	if( a.length == 2 && !a[0].match(/\D/))
		theCount = parseInt(a[0]);
	if( theCount != null && theCount > 0){
			if( a[1] == txtYears.toLowerCase() || a[1] == txtYear.toLowerCase())
		{
			dx.setFullYear(dx.getFullYear() + theCount);
			return dx.format(jelloDateFormatString());//dx.toDateString();
		}
		if( a[1] == txtMonths.toLowerCase() || a[1] == txtMonth.toLowerCase())
		{
			dx.setMonth(dx.getMonth() + theCount);
			return dx.format(jelloDateFormatString());//dx.toDateString();
		}
		if(a[1] == txtWeeks.toLowerCase() || a[1] == txtWeek.toLowerCase())
			theCount *= 7;
		dx.setDate(dx.getDate() + theCount);
		return dx.format(jelloDateFormatString(true));//dx.toDateString();
	}
	return null;
}
var txtWeeks="Weeks";
var txtDays="Days";
var txtMonths="Months";
var txtFrom="from";
var txtIn="in";
var txtWork="work";
var txtMobile="mobile";
var txtFax="fax";
var txtCompany="company";
var txtTitle="title";
var txtAddress="address";
var txtHomeAddress="homeaddress";
function isDateNoise(arg, words)
{
	if( arg == "")
		return false;
	// get rid of starting and trailing spaces
	arg = arg.replace(/^ */,"");
	arg = arg.replace(/ *$/,"");
	// all digits, hold it til we get more
	if(!arg.match(/\D/))
		return true;
	alow = arg.toLowerCase();
	if (alow == txtNext.toLowerCase() || alow == txtToday.toLowerCase() ||
			alow == txtTomorrow.toLowerCase() || alow == txtFrom.toLowerCase())
		return true;
	else if( alow == txtWeeks.toLowerCase() || alow == txtWeek.toLowerCase() || alow ==txtMonths.toLowerCase() ||
				alow == txtDays.toLowerCase() || alow == txtDay.toLowerCase() || alow == txtMonth.toLowerCase() ||
    		alow == txtYears.toLowerCase() || alow == txtYear.toLowerCase() ){
			return true;
		}
	else
		return false;
}


function isEmail(str) {

    if (str == null || str.length == 0) {
       return false;
    } else if (!emailAtSign( str )) {
       return false;
    } else if (!emailBracket(str)) {
       return false;
    }else if (badEmailPeriod(str)) {
       return false;
    } else if (!emailSuffix(str)) {
        return false;
    }
	return true;
}


function emailAtSign (addr) {
    // CHECK THAT THERE IS AN '@' CHARACTER IN THE STRING and only 1
    if (addr.indexOf ('@', 0) == -1) {
        return ( false );
	} else if(addr.indexOf('@') != addr.lastIndexOf('@')){
		return (false);
    }else if ( addr.indexOf ( '@', 0 ) < 1 ) { // CHECK THERE IS AT LEAST ONE CHARACTER BEFORE THE '@' CHARACTER
        return ( false );
    } else {
        return ( true );
    }
}


function emailBracket (addr) {
    if ( addr.indexOf ( '[', 0 ) == -1 && addr.charAt (addr.length - 1) == ']') {
        return ( false );
    } else  if (addr.indexOf ( '[', 0 ) > -1 && addr.charAt (addr.length - 1) != ']') {
        return ( false);
    } else {
        return ( true );
    }
}



function badEmailPeriod (addr) {
    if (addr.indexOf ( '@', 0 ) > 1 && addr.charAt (addr.length - 1 ) == ']')
        return ( false );
    if (addr.indexOf ( '.', 0 ) == -1)
        return ( true );

    return ( false );
}

function emailSuffix(addr) {

    if (addr.indexOf('@', 0) > 1 && addr.charAt(addr.length - 1) == ']') {
        return ( true );
    }

    var len = addr.length;
    var pos = addr.lastIndexOf ( '.', len - 1 ) + 1;
    if ( ( len - pos ) < 2 || ( len - pos ) > 4 ) {
        return ( false );
    } else {
        return ( true );
    }
}


// Declaring required variables
var digits = "0123456789";
// non-digit characters which are allowed in phone numbers
var phoneNumberDelimiters = "()- ";
// characters which are allowed in international phone numbers
// (a leading + is OK)
var validWorldPhoneChars = phoneNumberDelimiters + "+";
// Minimum no of digits in an international phone no.
var minDigitsInIPhoneNumber = 10;



function stripCharsInBag(s, bag)
{   var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++)
    {
        // Check that current character isn't whitespace.
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function isPhoneNumber(str){

	var bracket=3;

	if( str == null || str == "")
		return false;
	str = str.replace(/ /g,"");

	if(str.indexOf("+")>1) return false;
	if(str.indexOf("-")!=-1)bracket=bracket+1;
	if(str.indexOf("(")!=-1 && str.indexOf("(")>bracket)
		return false;
	var brchr=str.indexOf("(");
	if(str.indexOf("(")!=-1 && str.charAt(brchr+2)!=")")
		return false;
	if(str.indexOf("(")== -1 && str.indexOf(")")!=-1)
		return false;
	var s = stripCharsInBag(str,validWorldPhoneChars);
	// was using isDigit
	dcheck = /\D/;
	return (!dcheck.test(s) && s.length >= minDigitsInIPhoneNumber);
}
function endDateAfterStart(strt,end)
{
		var dx = new Date.parseDate(strt,jelloDateFormatString());
		var dy = new Date.parseDate(end,jelloDateFormatString());;
		if( dy < dx)
			return true;
		else
			return false;
}

function importText(confirmMode, passedText, passedCategory)
{
//get bulk text
//on confirm mode display stats and return confirmation dialog choice
var passedC = "";
var passedLine = false;
if( importText.arguments.length < 2 )
	var ct=Ext.getCmp("collectArea").getValue();
else{
	ct = passedText;
	if( typeof(ct) == "undefined"  || ct == null || ct == "")
		return;
	passedLine = true;
	if( typeof(passedCategory) != "undefined"  && passedCategory != null
			&& passedCategory != "")
		passedC = passedCategory;

}
var icount=0;
if (notEmpty(ct)==false){alert(txtMsgImportNothing);return;}
//split lines
var t=ct.split("\n");
var lline=t.length;
var vitems=new Array();
var newTags=new Array();
var invalids=new Array();
 //loop through lines
  for (var x=0;x<lline;x++)
  {
    var l=t[x];
    var o=l.charCodeAt(l.length-1);
    if (o==13){l=t[x].substr(0,t[x].length-1);}
    var it=l.split(",");
    if(notEmpty(it[0]))
    {
        vitems.push(it);

        if (it.length>1)
        {
        //loop through line's elements
        for (var y=1;y<it.length;y++)
        {
          var el=Trim(it[y]);
          if (notEmpty(el))
          {
            var r=getTag(el);
            tagStore.filter("tag",el);
				if (tagStore.getCount()==0 && isDate(el)==false)
				{
					if (!arrayExists(el,newTags)){newTags.push(el);}
				}
            tagStore.clearFilter();
          }
        }
        }
    }
    else
    {//if blank before first comma mark invalid
    if (l.length>1){invalids.push(x);}
    }
  }

  if (confirmMode)
  {
  //confirmation import message
  var nt="";
    for (var x=0;x<newTags.length;x++)
    {
    nt+=newTags[x]+" ";
    }
    var ms=txtMsgImportConfirmSimple.replace("%1",vitems.length);
    ms=ms.replace("%2",newTags.length);
    ms=ms.replace("%3",nt);
    ms=ms.replace("%4",invalids.length);
  return confirm(ms);
  }
  else
  {
  //import and convert lines to outlook items
  //first add new tags
    if (newTags.length>0)
    {
      for (var x=0;x<newTags.length;x++)
      {createTag(newTags[x]);}
    }
  //add items
    if (vitems.length>0)
    {
      for (var x=0;x<vitems.length;x++)
      {
         var vi=vitems[x];
         var cats="";
             var newA=createActionOL(vi[0]);
             if (vi.length>1)
             {
                for (var y=1;y<vi.length;y++)
                {
                var el=vi[y];
                  if (isDate(el)){newA.itemProperties.item(dueProperty).Value=el;}
                  else
                  {cats+=el+";";}
                }
             }
              newA.Categories= cats;
			var cln = "";
			if( passedC != "" && cats != "" && cats != null)
				cln += ";";
			newA.Categories += cln + passedC;
              newA.Save();
              updateTheLatestThing();
              icount++;
     }
    }
  if( !passedLine)
	Ext.getCmp("collectArea").setValue("");
var cmsg=txtAlColSucc1;
cmsg=cmsg.replace("%1",icount).replace("%2","<a class=jellolinktop onclick=pInbox();><b>Inbox</b></a>");
Ext.info.msg(txtCompleted,cmsg);
Ext.get("topInbox").highlight("#ff0000", {attr: "background-color",endColor: "ffffff",easing: 'easeIn',duration: 2});
if (icount>0 && !passedLine){Ext.getCmp("cinb").setText(txtInbox+"&nbsp;("+countInboxItems()+")");Ext.getCmp("cinb").getEl().highlight("#ff0000", {attr: "background-color",endColor: "ffffff",easing: 'easeIn',duration: 2});}
  }

}

//*** end new collect code


function lastNonBlankLine(t)
{
//return the last non blank line
var y=(t.length)-1;var ll=y;
  for (var x=y;x>0;x--)
  {
  if (t[x].length>1){ll=x;break;}
  }
  return ll;
}

function checkNewTags(addTags)
{
//check user entered tags and create non existing
if (notEmpty(addTags)==false){return 0;}
var newTags=new Array();

var it=addTags.split(",");

        for (var y=0;y<it.length;y++)
        {
          var el=Trim(it[y]);
          if (notEmpty(el))
          {
            var r=getTag(el);


            tagStore.filter("tag",el);
            if (tagStore.getCount()==0){newTags.push(el);}
            tagStore.clearFilter();
          }
        }

//add the tags
    var nt=0;
    if (newTags.length>0)
    {
      for (var x=0;x<newTags.length;x++)
      {createTag(newTags[x]);nt++;}
    }
return nt;
}

function saveColDef()
{
//save collect definition
if (thisDefinition>0)
{
var choice=confirm(txtColPrmReplDef);if (!choice){return;}
delSavedDef(thisDefinition,true);
}
var fid=Ext.getCmp("folid").getValue();
var arc=Ext.getCmp("marc").getValue();//archive
var rem=Ext.getCmp("mrem").getValue();//remove original task
var cre=Ext.getCmp("mcre").getValue();//create non existing tags
var cim=Ext.getCmp("cas").getValue();//tag newly created actions with this list of tags
var dim=Ext.getCmp("mdim").getValue();//do not import existing categories of items
var don=Ext.getCmp("mdon").getValue();//ignore done tasks
var dup=Ext.getCmp("mdup").getValue();//replace duplicates
var sfil=Ext.getCmp("sfil").getValue();
var cfil=Ext.getCmp("cfil").getValue();
var filt=iDasl;
try{
var nm=NSpace.GetFolderFromID(fid);
}catch(e){alert(txtColNoSave);return;}

jello.lastImportSetId++;

var tr=new impSetRecord({
id:jello.lastImportSetId,
name:nm.FolderPath,
folderid:fid,
filter:filt,
removeoriginal:rem,
createtags:cre,
tagnew:cim,
noexistingcats:dim,
ignoredone:don,
replacedups:dup,
archive:arc,
sfilter:sfil,
cfilter:cfil
});
impSetStore.add(tr);
syncStore(impSetStore,"jello.importSettings");
jese.saveCurrent();

alert(txtColSaved);
thisDefinition=jello.lastImportSetId;
getSavedDefs();
}

function getSavedDefs()
{//refresh bottom panel with saved import definitions
var ret="<b>"+txtColSvdDefsLst+":</b><br>";

  var toList = function(rec){
  ret+="<img src=img\\olpreview.gif style=float:left;><a title='Load this definition' class=jellolink onclick=runSavedDef("+rec.get("id")+",true)>"+rec.get("name")+"</a>&nbsp;<a class=jellolink onclick=runSavedDef("+rec.get("id")+") title='Run this definition'><span class=smallfunction>["+txtColSvdDefRun+"]</span></a>&nbsp;<a class=jellolink onclick=renameSavedDef("+rec.get("id")+") title='Rename this definition'><span class=smallfunction>["+txtRename+"]</span></a>&nbsp;<a class=jellolink onclick=delSavedDef("+rec.get("id")+") title='Remove this definition'><span class=smallfunction>["+txtDelete+"]</span></a><br>";
  };

impSetStore.each(toList);
return ret;
}

function delSavedDef(id,noalert)
{
//delete a saved definition
var ix=impSetStore.find("id",new RegExp("^"+id+"$"));
var r=impSetStore.getAt(ix);
var tnr=r.get("name");
if (noalert==null || noalert==false)
{if (confirm(txtColSvdDefsRmQ+":"+tnr+" ?")==false){return;}}
impSetStore.remove(r);
syncStore(impSetStore,"jello.importSettings");
jese.saveCurrent();
getSavedDefs();
}

function renameSavedDef(id)
{
//rename a saved definition
var ix=impSetStore.find("id",new RegExp("^"+id+"$"));
var r=impSetStore.getAt(ix);
var tnr=r.get("name");
var newname=prompt(txtColSvdDefsNam,tnr);
if (!notEmpty(newname)){return;}
r.beginEdit();
r.set("name",newname);
r.endEdit();

syncStore(impSetStore,"jello.importSettings");
jese.saveCurrent();
getSavedDefs();
}
function runSavedDef(id,editor)
{
//open a saved definition
var ix=impSetStore.find("id",new RegExp("^"+id+"$"));
var r=impSetStore.getAt(ix);
try{
var fid=r.get("folderid");var fn=r.get("name");
var fd=NSpace.GetFolderFromID(fid);
}catch(e){alert(txtColPrmNonExFld);return;}
var arc=r.get("archive");
var rem=r.get("removeoriginal");
var cre=r.get("createtags");
var cim=r.get("tagnew");
var dim=r.get("noexistingcats");
var don=r.get("ignoredone");
var dup=r.get("replacedups");
var flt=r.get("filter");
var sflt=r.get("sfilter");
var cflt=r.get("cfilter");

    if (editor==true)
    {//open in editor
    collectOLFolder(fid,fd,arc,rem,cre,cim,dim,don,dup,flt,sflt,cflt);
    thisDefinition=id;
    updateColStatus();
    }
    else
    {//run import definition
      try{
        if(notEmpty(flt)){var tc=fd.Items.Restrict("@SQL="+flt);}else{var tc=fd.Items;}
      }catch(e){alert(txtColPrmDefEr);return;}
    importOLFolder(fid,flt,arc,rem,cre,cim,dim,don,dup);
    }
}

function getPreviewTitlesFromDef(tc)
{
//get first items from definition execution
var cnt=tc.Count;var ret="";
if (cnt>5){cnt=5;}
if (cnt==0){return "*No Items*";}
  for (var x=1;x<=cnt;x++)
  {
  ret+=tc(x)+"\n";
  }
  return ret;
}

function checkDuplicates(it,itDate,itParent)
{
//check item for duplicates in the same folder.
//If duplicate is found the one last changed is kept
var fld=it.Parent.Items;
var f=fld.Restrict("@SQL=urn:schemas:httpmail:subject='"+it.Subject+"' AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2");
  if (f.Count>1)
  {
  //duplicate found.
  //check which item was modified last and keep it
  var oldItem;var newItem;var oParent="";var nParent="";

	 f.Sort("[DAV:getlastmodified]",true);
		var checkedItem=0;var otherItem=0;
		for (var x=1;x<=f.Count;x++)
		{//loop through items. Find original item and one more
			if (it.EntryID==f(x).EntryID){checkedItem=x;}
			else{if (otherItem==0){otherItem=x;}}
		}
		var dChecked=itDate.clone();var dOther=new Date(f(otherItem).LastModificationTime);
		if (dChecked<dOther){var olderItem=f(checkedItem);oParent=itParent;var ddo=dChecked.clone();var newerItem=f(otherItem);nParent=f(otherItem).Parent;var ddn=dOther.clone();}
		else{var newerItem=f(checkedItem);nParent=itParent;var ddn=dChecked.clone();var olderItem=f(otherItem);oParent=f(otherItem).Parent;var ddo=dOther.clone();}
		var smsg=txtColAlDupDisc;
    if (nParent.EntryID==jello.actionFolder)
    {
    smsg=smsg.replace("%0",txtDiscard).replace("%1",oParent.FolderPath+"\\"+olderItem).replace("%2",DisplayDate(ddo)+"@"+DisplayTime(ddo)).replace("%3",nParent.FolderPath+"\\"+newerItem).replace("%4",DisplayDate(ddn)+"@"+DisplayTime(ddn));
    var choice=confirm(smsg);}
    else
    {
        smsg=smsg.replace("%0",txtReplace).replace("%1",oParent.FolderPath+"\\"+olderItem).replace("%2",DisplayDate(ddo)+"@"+DisplayTime(ddo)).replace("%3",nParent.FolderPath+"\\"+newerItem).replace("%4",DisplayDate(ddn)+"@"+DisplayTime(ddn));
    var choice=confirm(smsg);}

    if (choice==true){olderItem.Delete();}

  }
}

function collectHelp()
{
          Ext.Msg.show({
           title:txtHelp,
           msg:txtCollectInfo+txtCollectInfoAdv ,
           buttons: Ext.Msg.OK,
           animEl: 'elId',
           icon: Ext.MessageBox.INFORMATION
        });
}

function importTextAdvanced(confirmMode, passedText, passedCategory)
{
	//get bulk text
	//on confirm mode display stats and return confirmation dialog choice
	var ct;
	var passedC = "";
	var passedLine = false;
	if( importTextAdvanced.arguments.length < 2 )
		ct=Ext.getCmp("collectArea").getValue();
	else{
		ct = passedText;
		if( typeof(ct) == "undefined"  || ct == null || ct == "")
			return;
		passedLine = true;
		if( typeof(passedCategory) != "undefined"  && passedCategory != null
				&& passedCategory != "")
			passedC = passedCategory;

	}
	var icount=0;var coCount=0;var caCount=0;var nCount=0;
	if (notEmpty(ct)==false){alert(txtMsgImportNothing);return;}
	//split lines
	var t=ct.split("\n");
	var lline=t.length;
	var vitems=new Array();
	var newTags=new Array();
	var invalids=new Array();
	 //loop through lines
   	var lastLine = "";
	for (var x=0;x<lline;x++)
	{
	    var l=t[x];
	    var o=l.charCodeAt(l.length-1);
	    if (o==13){l=t[x].substr(0,t[x].length-1);}
		l = l.replace(/ +/g," ");
		l = Trim(l);
		if( l == null || l == "")
			continue;
			// continuation char ar end of line?
		o = l.charAt(l.length -1);
		if( o == '\\' )
		{
			if( lastLine != "")
				lastLine += " ";
			lastLine += l.substring(0,l.length-1);
			// is this the last line?
			if( x+1 < lline)
				continue;
			l = "";
		}
		if( lastLine != "")
		{
			l = lastLine + " " + l;
		}
		lastLine = "";
	    var it=l.split(" ");

		if( it.length == 0 ) continue;
		var xline = new LineObject();
		// first element has the xxx: type specification
		// of is just the first word of a task
		xline.checkType(it[0]);

		//loop through line's elements
		for (var y=1;y<it.length;y++)
		{
			var el=Trim(it[y]);
			if (!notEmpty(el))
				continue;
			var res = xline.nextWord(el,newTags);

		}
		if( cDebug) alert(xline.toString());
		// push this line's onject on the stack
		vitems.push(xline);

	}



  if (confirmMode)
  {
  //confirmation import message
  var nt="";
    for (var x=0;x<newTags.length;x++)
    {
    nt+=newTags[x]+" ";
    }

	var cp=0;var ca=0;var cn=0;var cc=0;
	for (var x=0;x<vitems.length;x++)
	{
		cc++;
		if (vitems[x].type=="p"){cp++;cc--;}
		if (vitems[x].type=="a"){ca++;cc--;}
		if (vitems[x].type=="n"){cn++;cc--;}
	}

	var countTxt="";
	if (cc>0){countTxt+=cc+" Actions\n";}
	if (cp>0){countTxt+=cp+" Contacts\n";}
	if (ca>0){countTxt+=ca+" Appointments\n";}
	if (cn>0){countTxt+=cn+" Notes\n";}

    var ms=txtMsgImportConfirm.replace("%1",countTxt);
    ms=ms.replace("%2",newTags.length);
    ms=ms.replace("%3",nt);
    ms=ms.replace("%4",invalids.length);
  return confirm(ms);
  }
  else
  {
  //import and convert lines to outlook items
  //first add new tags
    if (newTags.length>0)
    {
      for (var x=0;x<newTags.length;x++)
      {createTag(newTags[x]);}
    }
  //get UD folders
  var jcalendarItems=setAndCheckArcFolder(jello.calendarFolder);
  var jcontactItems=setAndCheckArcFolder(jello.contactFolder);
  if (jcalendarItems==null){jcalendarItems=calendarItems;}
  if (jcontactItems==null){jcontactItems=contactItems;}
  //add items
    if (vitems.length>0)
    {
      for (var x=0;x<vitems.length;x++)
      {
         var vi=vitems[x];
			// lets see if the user specified an item type to create type to create
			var newA;
			if( vi.type == "a"){
				newA = jcalendarItems.Items.Add();
				newA.Subject = vi.desc;caCount++;
			}else if( vi.type == "n"){
				newA = noteItems.Items.Add();
				newA.Body = vi.desc;nCount++;
			}else if (vi.type == "p" && vi.desc.length != 0){
				newA = jcontactItems.Items.Add();coCount++;
			}else
				{
				newA=createActionOL(vi.desc);icount++;
				}

				var jfmt = jelloDateFormatString();
				var jfmty = jfmt.indexOf('y');
				if( jfmty >= 0)
						jfmt = jfmt.replace('y','Y');
             var dx = new Date().format(jfmt);//('Y/m/d');
			 // no dates than make the dates today
	if(vi.type == "a" && (vi.startDate == "" && vi.endDate == "")){
						vi.startDate =  vi.endDate =dx;
			}
			// a task
			if( vi.type == "t"){
				// just one date, make it end and today start
				if( vi.startDate != "" && vi.endDate == "")
				{
					vi.endDate = vi.startDate;
					vi.startDate = dx;
				}else if( vi.startDate == "")
					vi.startDate = vi.endDate;
				if( endDateAfterStart(vi.startDate, vi.endDate)){
					Ext.Msg.alert(txtColBadEDateT, txtColBadEDate + vi.desc );
					newA.Delete();
					continue;
				}

				if (newA.Class==48){
					if( vi.endDate != null && vi.endDate != "")
						newA.ItemProperties.Item(dueProperty).Value=new Date.parseDate(vi.endDate,jfmt).format('Y/m/d');
				    if( vi.startDate != null && vi.startDate != "")
						newA.ItemProperties.Item("StartDate").Value=new Date.parseDate(vi.startDate,jfmt).format('Y/m/d');
				}

				if (newA.Class==26){
        newA.ItemProperties.Item("Start").Value=new Date.parseDate(vi.startDate,jfmt).format('Y/m/d');
				newA.ItemProperties.Item("End").Value=new Date.parseDate(vi.endDate,jfmt).format('Y/m/d');
				}


			}else if( vi.type == "a"){
        if( vi.startTime == "" && !vi.allDay){
					Ext.Msg.alert(txtColBadETimeT,txtColBadETime + vi.desc );
					newA.Delete();
					continue;
				}
				if( vi.endDate == "" )
						vi.endDate = vi.startDate;
        if( endDateAfterStart(vi.startDate, vi.endDate) && !vi.allDay){
					Ext.Msg.alert(txtColBadEDateT, txtColBadEDate + vi.desc );
					newA.Delete();
					continue;
				}
				if( vi.allDay == true){

					newA.itemProperties.item("Start").Value = Date.parseDate(vi.startDate,jfmt).format('Y/m/d') + " 12:00AM";
					newA.itemProperties.item("AllDayEvent").Value = true;
	        var xdt = new Date.parseDate(vi.endDate,jfmt);
					xdt.setDate(xdt.getDate()+1);
					vi.endDate = xdt.format(jfmt);
					//vi.endDate = vi.endDate.substring(vi.endDate.indexOf(" ")+1);
					newA.itemProperties.item("End").Value = Date.parseDate(vi.endDate,jfmt).format('Y/m/d');
					newA.itemProperties.item("BusyStatus").Value = 0;

				}else{
					newA.itemProperties.item("Start").Value = Date.parseDate(vi.startDate,jfmt).format('Y/m/d') + " " + vi.startTime;
					if( vi.endTime == "")
						newA.itemProperties.item("Duration").Value = 60;
					else
						newA.itemProperties.item("End").Value = Date.parseDate(vi.endDate,jfmt).format('Y/m/d') + " " + vi.endTime ;
					newA.itemProperties.item("BusyStatus").Value = 2;
				}
			}else if( vi.type == "p"){
				var des = Trim(vi.desc);
				if( des == "" ){
					Ext.Msg.alert(txtColBadPersT,txtColBadPers + vi.toString(true) );
					newA.Delete();
					continue;
				}
				var xdesc = vi.desc.split(" ");
				if( xdesc.length >= 2){
					newA.itemProperties.item(69).Value = xdesc[0];
					newA.itemProperties.item(97).Value = xdesc[1];
				}else
					newA.itemProperties.item(97).Value = xdesc[0];
				if( xdesc.length > 2)
					for( var i=2; i < xdesc.length; i++)
						newA.itemProperties.item("Body").Value += xdesc[i];
				for( var i=0; i < vi.email.length; i++){
					var mtype = vi.email[i][0];
					if( mtype == txtWork.toLowerCase())
						newA.itemProperties.item(56).Value = vi.email[i][1];
					if( mtype == txtHome.toLowerCase())
						newA.itemProperties.item(60).Value = vi.email[i][1];
					else
						newA.itemProperties.item(64).Value = vi.email[i][1];
				}
				newA.itemProperties.item(52).Value = vi.company;
				newA.itemProperties.item(89).Value = vi.title;
				newA.itemProperties.item(35).Value = vi.address;
				newA.itemProperties.item(77).Value = vi.homeAddress;
				for( var i=0; i < vi.phone.length; i++){
					var mtype = vi.phone[i][0];
					if( mtype == txtWork.toLowerCase())
						newA.itemProperties.item(44).Value = vi.phone[i][1];
					else if( mtype == txtHome.toLowerCase())
						newA.itemProperties.item(85).Value = vi.phone[i][1];
					else if( mtype == txtMobile.toLowerCase())
						newA.itemProperties.item(108).Value = vi.phone[i][1];
					else if( mtype == txtFax.toLowerCase())
						newA.itemProperties.item(42).Value = vi.phone[i][1];
				}

			}

			newA.Categories= vi.cats;
			var cln = "";
			if( passedC != "" && vi.cats != "" && vi.cats != null)
				cln += ";";
			newA.Categories += cln + passedC;
			if( vi.body.length != 0)
				newA.itemProperties.item("Body").Value = newA.itemProperties.item("Body").Value + " " + vi.body;
			newA.Save();
			updateTheLatestThing();
			//icount++;
     }
    }
	if( !passedLine)
		Ext.getCmp("collectArea").setValue("");
    var cmsg=txtAlColSucc2;
    var counters="";
    if (icount>0){counters+=icount+" "+txtCbActions;}
    if (coCount>0){counters+="<br>"+coCount+" "+txtContacts;}
    if (caCount>0){counters+="<br>"+caCount+" Appointments";}
    if (nCount>0){counters+="<br>"+nCount+" Notes";}
	cmsg=cmsg.replace("%1",counters).replace("%2","<a class=jellolinktop onclick=pInbox();><b>Inbox</b></a>");
  Ext.info.msg(txtCompleted,cmsg);
  Ext.get("topInbox").highlight("#ff0000", {attr: "background-color",endColor: "ffffff",easing: 'easeIn',duration: 2});
if (icount>0 && !passedLine){Ext.getCmp("cinb").setText(txtInbox+"&nbsp;("+countInboxItems()+")");Ext.getCmp("cinb").getEl().highlight("#ff0000", {attr: "background-color",endColor: "ffffff",easing: 'easeIn',duration: 2});}

  }

}

function toggleAdvancedMode()
{
jello.collectAdvancedMode=!jello.collectAdvancedMode;
if (jello.collectAdvancedMode){Ext.getCmp("cjoin").show();}else{Ext.getCmp("cjoin").hide();}
jese.saveCurrent();
}

function doTextCollect()
{//run the text collect process
	Ext.getCmp("cotags").focus();
	if(jello.collectAdvancedMode==true || jello.collectAdvancedMode!=0)
	{var c=importTextAdvanced(true);
	if (c==true){importTextAdvanced(false);}}
	else
	{var c=importText(true);
	if (c==true){importText(false);}}
	setTimeout(function(){Ext.getCmp("cotags").reset();Ext.getCmp("collectArea").focus();},3);
}

function doJoinSelected()
{
	var w = Ext.getCmp("collectArea");
	if( w == null)
		return;
	var t = w.getEl().dom.document.selection.createRange().text;
	var tnew = t.replace(/\r\n/g, " \\\n");
	var rtext = w.getRawValue();
	var pos = rtext.indexOf(t);
	rtext = rtext.substring(0,pos) + tnew + rtext.substring(pos+t.length);
	w.setRawValue(rtext);
}