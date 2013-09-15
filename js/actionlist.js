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


var isProject=false;
//define stores
var priorities = [
        ['99',txtNoPriority],
        ['1',txtPriority+' 1 ('+txtHigh+')'],
        ['2',txtPriority+' 2'],
        ['3',txtPriority+' 3 ('+txtLow+')']

        ];

var filters=[
        ['0',txtUncompleted],
        ['1',txtCompleted],
        ['5',txtNextActions],
        ['3',txtNotStarted],
        ['2',txtInProgress],
        ['6',txtRecAdded],
        ['7',txtRecCompleted],
        ['8',txtStalled],
        ['4',txtAllActions]
        ];

var filterValues=new Ext.data.SimpleStore({
        fields: [
           {name: 'value'},
           {name: 'text'}
        ],
        data:filters
    });


var priorityStore=new Ext.data.SimpleStore({
        fields: [
           {name: 'value'},
           {name: 'text'}
        ],
        data:priorities
    });

var columnStore=new Ext.data.JsonStore({
        fields: [
           {name: 'grid'},
           {name: 'data'}
        ],
        data:jello.gridColumns
    });

var columnRecord = Ext.data.Record.create([
    {name: 'grid'},
    {name: 'data'}
]);

//define actionRecord
 var actionRecord = Ext.data.Record.create([
    {name: 'subject'},
    {name: 'type'},
    {name: 'entryID'},
	{name: 'importance'},
	{name: 'attachment'},
	{name: 'created',type:'date'},
	{name: 'iclass'},
	{name: 'to'},
	{name: 'cc'},
	{name: 'contacts'},
	{name: 'sensitivity'},
	{name: 'sortDate'},
	{name: 'status'},
	{name: 'due',type:'date'},
	{name: 'icon'},
	{name: 'body'},
	{name: 'notes'},
	{name: 'unread',type:'boolean'},
	{name: 'attachmentCount',type:'int'},
	{name: 'attachmentList'},
	{name: 'tags'},
	{name:'markfordownload',type:'int'},
	{name:'completed',type:'date'},
	{name:'duration',type:'int'},
	{name: 'flag',type:'int'},
  // added for tables use in OL12 and above
	{name: 'groupon'}


]);

function showContext(tid,fun,dasl,filtering,oItems,istag,returnItems){
//Query tasks for a context's item
//oItems are other types (folders and tags) which will display too
//fun=0 count only, 1 show list, 2 return store
if (typeof(tid)=="number"){var ctx=getTagName(tid);}
else{var ctx=tid;tid=getTagID(ctx);}
	if (filtering==null)
	{
	//no filtering defined. use last saved filter for tag
	filtering=getTagFilter(tid);
	}
ctx=ctx.replace(new RegExp("~","g")," ");
//try{var ctx=getTagName(tid);}catch(e){ctx="Jello Objects";}
if (istag==false){buildGrid(ctx,null,0,filtering,oItems,istag);}
if (istag==null){istag=true;}
thisViewType="action";
if (dasl=="undefined"){dasl=null;}
if (filtering=="undefined" || filtering==null){filtering=0;}
var iF;var counter=0;var pCount=0;var tCount=0;
//take care and check context string
var pCtx=ctx;
if (fun==1 || fun==2){window.status=txtFetching;}else{window.status=txtCounting;}
//set DASL filter
if (dasl==null){dasl="(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "')";}
//set filtering
if (filtering=="0"){//uncompleted
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
}
if (filtering=="1"){//completed
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 = 2";
}
if (filtering=="2"){//started
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 = 1";
}
if (filtering=="3"){//not started
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 = 0";
}
if (filtering=="5"){//NAs
var SNextAction=getActionTagName();
dasl+=" AND (urn:schemas-microsoft-com:office:office#Keywords LIKE '" + SNextAction + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + SNextAction + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + SNextAction + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + SNextAction + "')";
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
}
if (filtering=="6"){//Recently added
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
dasl+=" AND %last7days(urn:schemas:calendar:created)%";
}
if (filtering=="7"){//Recently completed
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 = 2";
dasl+=" AND %last7days(http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/810f0040)%";
}
if (filtering=="8"){//Stalled actions created older than 30 days
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
//dasl+=" AND %lastmonth(urn:schemas:calendar:created)%";
var fDate=new Date();
fDate.setDate(fDate.getDate()-30);
var STstart=fDate.getUTCFullYear() + "/" + (fDate.getUTCMonth()+1)+"/"+fDate.getUTCDate()+" "+fDate.getUTCHours()+":"+fDate.getUTCMinutes();
dasl+=" AND (urn:schemas:calendar:created<'"+STstart+"')";
}
if (filtering=="9"){//waiting
var SWaitAction=getWaitTagName();
dasl+=" AND (urn:schemas-microsoft-com:office:office#Keywords LIKE '" + SWaitAction + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + SWaitAction + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + SWaitAction + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + SWaitAction + "')";
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
}

if (filtering=="99"){//2nd tag
dasl+=" AND (urn:schemas-microsoft-com:office:office#Keywords LIKE '" + lastContext + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + lastContext + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + lastContext + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + lastContext + "')";
dasl+=" AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 0";
}

if (filtering=="4"){//all
//no filters
}

var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;

var its=iF.Restrict("@SQL="+dasl);

counter=its.Count;


	if (fun==0)
  {
	//counter only
	window.status="";
   if (returnItems){return its;}else{return counter+"";}
   }
	else
	{

	itmContent=pCtx;
	if (filtering!="99" && ctx!="##"+txtInbox){lastContext=ctx;}

	if (oItems==null)
	{
	//if this is a tag get subtags to show too


    //if (tid==-1){alert("Error");return;}

    if (tid>0)
    {
    var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
    var r=tagStore.getAt(ix);
    var ftag=r.get("istag");

      if (ftag==true)
      {
      var oItems=new Array();
      tagStore.filter("parent",new RegExp("^"+tid+"$"));
      for (var x=0;x<tagStore.getCount();x++)
        {
        var r=tagStore.getAt(x);
        oItems.push(r);
        }
        tagStore.clearFilter();
        lastOpenTagID=tid;
      }
  }
	}
status=txtReady;
	if (fun==1){buildGrid(ctx,its,counter,filtering,oItems,istag);return;}
	if (fun==2)
  {
  var theNewStore= buildGrid(ctx,its,counter,filtering,oItems,istag,true);
  fun=0;
  return theNewStore;
  }

	}
}

function editAction(actionItems,noGrid,assignLast,assignTag){
//open action edit form

	//It shows after adding another tag
getTagsArray();
try{var t=Ext.getCmp("theactionform");t.destroy();}catch(e){}
thisViewType="actionform";
try{var itcount=actionItems.length;}catch(e){itcount=0;};
if (itcount>0){if(actionItems[0].get("type")=="t"){editTag(actionItems);return;}}

var fsubj=txtMultiple;
var tgs="";
var dts=jello.dateSeperator;
var dccformat="n"+dts+"j"+dts+"Y";
var dccalt="m/d|n/j|m-d|n-j|m/d/y|m/d/y|m/d/y|n/j/y|j/n/y";
if(jello.dateFormat == 0 || jello.dateFormat == "0"){dccformat="j"+dts+"n"+dts+"y";dccalt="d/m|j/n|d-m|j-n|d/m/y|d/m/y";}
if(jello.dateFormat == 2 || jello.dateFormat == "2"){dccformat="Y"+dts+"m"+dts+"d";dccalt="d/m|n/j|n-j|y/m/d|y/m/d/Y-m-d";}

var ftagList="";
var fdue=null;
var fnotes=null;fnotesavail=true;
var fpriority=99;
var finfo="";
var folid=null;
var att="";
var tpbar=[];
var totlwork="";

if (itcount==0){
//handle new item
	fsubj=txtNewAction;
	//if (assignLast!=false){ftagList=lastContext;tgs=tagList(ftagList,ftagList,ftagList);}else{}
	//if (notEmpty(assignTag)){ftagList=assignTag;tgs=tagList(ftagList,ftagList,ftagList);}
	if (assignLast!=false){ftagList=lastContext+";";tgs=tagList(null,null,ftagList);}else{}
	if (notEmpty(assignTag)){ftagList=assignTag;tgs=tagList(null,null,ftagList);}

folid="000";
finfo="<span class=tagsystem>"+txtNewAction+"</span>";
}

if (itcount==1){
//handle one item
try{var it=NSpace.GetItemFromID(actionItems[0].get("entryID"));} catch(e){return;}

	folid=it.EntryID;
	fsubj=it.Subject;
	if( it.Class == 43)
		var tt = new Date(getJDueProperty(it));
	else
	  try{var tt=new Date(it.itemProperties.item(dueProperty).Value);}catch(e){}
	fdue=tt;if (fdue.getYear()>4000){fdue=null;}
	tgs=tagList("formtagdisplay",null,it.itemProperties.item(catProperty).Value);
	ftagList=it.itemProperties.item(catProperty).Value;
	if(tgs==""){tgs=" ";}
	fnotes=getJNotesProperty(it);
	fpriority=getJPriorityProperty(it);
	var ct=DisplayDate(new Date(it.CreationTime));
	finfo="<nobr>"+txtCreated+": "+ct+"</nobr>&nbsp;&nbsp;";
	var ct=DisplayDate(new Date(it.LastModificationTime));
	finfo+="<nobr>"+txtModified+": "+ct+"&nbsp;&nbsp;</nobr>";
	att=getActionAttachmentLink(it);
	if (it.Complete==true){
  var ct=DisplayDate(new Date(it.DateCompleted));
	finfo+="<br><span class=tagarchive>"+txtCompleted+": "+ct+"</span></nobr>";}
	tpbar=getActionFormToolbar(finfo,true);
	if (OLversion>11){totlwork=it.TotalWork;}
}

if (itcount>1)
{
//handle multiple items
	for (x=0;x<actionItems.length;x++)
	{
	var it=NSpace.GetItemFromID(actionItems[x].get("entryID"));
	ftagList+=it.itemProperties.item(catProperty).Value+";";
	tgs="";
	fpriority=getJPriorityProperty(it);
    fnotesavail=false;
    	if( it.Class == 43)
			var tt = new Date(getJDueProperty(it));
		else
	   try{var tt=new Date(it.itemProperties.item(dueProperty).Value);}catch(e){}
	if (fdue==null){fdue=tt;}
	if (fdue.getYear()>4000){fdue=null;}
	}
	ftagList=getUniqueTags(ftagList);
	tgs=tagList("formtagdisplay",null,ftagList);
}

var notesStatus=txtOLBodyReplace;
var footerlabel = txtWebs+":<br><br><a class=jelloformlink onclick=actionEditContacts();>"+txtContacts+"</a>&nbsp;<span class=fkey>[F8]</span><br><a class=jelloformlink onclick=actionEditAttachments();>"+txtAttachFile+"</a>&nbsp;<span class=fkey>[F7]</span>";
if( folid != "000")
	footerlabel += "<br><a class=jelloformlink onclick=popTaskBody('"+it.EntryID+"',null);>"+txtBody+"</a>";

if (jello.autoUpdateTaskNotes==false){notesStatus=txtOLNoBodyReplace;}
 var simple = new Ext.FormPanel({
        labelWidth: 75, // label settings here cascade unless overridden
        frame:true,
        title: txtTaskName+' '+fsubj,
        width:580,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        iconCls:'actionformicon',
        id:'actform',
        buttonAlign:'center',
        defaults: {width: 400},
        defaultType: 'textfield',

        items: [{
                fieldLabel: txtSubject,
                id: 'subj',
                disabled:(txtMultiple==fsubj),
                value:fsubj,
                tabIndex:0,
                allowBlank:false
            },
                new Ext.form.ComboBox({
                fieldLabel: txtTags,
                id: 'tags',
                store:globalTags,
                hideTrigger:false,
                typeAhead:true,
                mode:'local',
                listeners:{
        blur:function(){quickAddTag("tags", "tagfield","formtagdisplay", Ext.getCmp("olid").getValue());},
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                //f.el.blur();
                quickAddTag("tags", "tagfield","formtagdisplay", Ext.getCmp("olid").getValue());}},

            select: function(cb,rec,idx){
            quickAddTag("tags", "tagfield","formtagdisplay", Ext.getCmp("olid").getValue());
            status=txtReady;
            }
		}

            }),
				new Ext.form.Label({
				html:tgs[0],
				height:40,
				cls:'formtaglist',
				id:'formtagdisplay'
				}),

				new Ext.form.Hidden({
				id:'tagfield',
				value:ftagList
				}),
     
         {xtype:'container',
          layout:'hbox',
          fieldLabel:txtDueDate,
          items:[ 
             new Ext.form.DateField({
                fieldLabel: txtDueDate,
                id: 'due',
                format:dccformat,
				invalidText:txtInvDate,
				width:100,
				altFormats:dccalt,
				listeners:{blur:function(ff){checkForSmartDate(ff);}},
                value:fdue
            })
            , new Ext.form.Label({
				html:'&nbsp;&nbsp;&nbsp;'+txtPriority+' :' ,
        width:60,
        style:{'margin-top':'3px'}
            }),
                new Ext.form.ComboBox({
                fieldLabel: txtPriority,
                id: 'pri',
                store:priorityStore,
                hideTrigger:false,
                selectOnFocus:true,
				editable:false,
                displayField:'text',
                valueField:'value',
                value:fpriority,
                width:80,
                triggerAction:'all',
                forceSelection:true,
                mode:'local'
                }),
                new Ext.form.Label({
				html:'&nbsp;&nbsp;Total work (mins) :' ,
        width:120,
        style:{'margin-top':'3px'}
            }),{
            id:'totlwork',
            xtype:'textfield',
            width:40,
            style:{'margin-top':'0px'},
            value:totlwork,
            hidden:OLversion<12
            }
                
                //,
               ]},
              
              
                new Ext.form.TextArea({
                fieldLabel: txtNotes,
                hidden:!fnotesavail,
                hideLabel:!fnotesavail,
                height:90,
                id: 'notes',
                emptyText:'(' + txtAcFrmNotes + ' ' + notesStatus + ')',
                value:fnotes,
                oldvalue:fnotes
            }),
            new Ext.form.Label({
				html:att,
				id:'actioncons',
				height:90,
				cls:'formattlist'
            }),
            new Ext.form.Label({
				html:footerlabel,
				hidden:!fnotesavail,
				cls:'formeditlist'
            }),

        new Ext.form.Hidden({
				id:'olid',
				value:folid
				}),
        new Ext.form.Hidden({
				id:'isSaved',
				value:null
				})
        ]
    });

var wttl=txtEdit;
  if (fsubj==txtNewAction){wttl=txtNewRecord;}


   var win = new Ext.Window({
        title: wttl,
        width: 600,
        tbar:tpbar,
        height:500,
        id:'theactionform',
        minWidth: 300,
        minHeight: 200,
        resizable:false,
        draggable:false,
        layout: 'fit',
        plain:true,
        draggable:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        modal:true,
        items: simple,
        listeners: {destroy:function(t){try{showOLViewCtl(true);}catch(e){}}},
        buttons: [{
            text:'<b>'+txtSave+'</b>',
            id:'sbut',
            tooltip:txtSave+' [F2]',
            listeners:{
            click: function(b,e){
            saveActionForm(b,noGrid);
            }
            }
        },{
            text:txtSave + ' & New',
            id:'snbut',
            tooltip:txtSave+' & New action',
            listeners:{
            click: function(b,e){
            saveActionForm(b,noGrid);
            editAction(null,true,false);
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

//show the form
    showOLViewCtl(false);
    win.show();
    win.setActive(true);
  setTimeout(function(){try{Ext.getCmp("subj").focus(true,100);}catch(e){}},1);


//shortcut keys
var actionKeyMap = new Ext.KeyMap('theactionform',
	[
	{
    key: Ext.EventObject.F2,
    stopEvent:true,
    fn: function(){
    try{saveActionForm(null,noGrid);}catch(e){}},
    scope: this},{
    key: Ext.EventObject.F8,
    stopEvent:true,
    fn: function(){
    actionEditContacts();},
    scope: this},{
    key: Ext.EventObject.F7,
    stopEvent:true,
    fn: function(){
    actionEditAttachments();},
    scope: this},{
    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
            try{win.destroy();}catch(e){}

            },
    scope: this}

    ]);

}

function saveActionForm(b,noGrid,returnID){
//j5
//action form save button handler

var g=getActiveGrid();
var ids=new Array();
var cl=0;
var xmsg="";

  if (returnID!=true)
  {
    if (noGrid!=true)
    {
    ids=getCheckedItems();
    cl=ids.length;
    var gs=g.getStore();
    }
    else
    {
    var cl=1;
    var olid=Ext.getCmp("olid").getValue();
    }

    if (notEmpty(Ext.getCmp("isSaved").getValue())){cl=1;}

    try{if(Ext.getCmp("olid").getValue()=="000"){cl=0;var ids=new Array();}}catch(e){}
  }
  else {var cl=0;
  try{var gs=g.getStore();}catch(e){}
  }
	//new action

	if (cl==0)
	{
	if(Ext.getCmp('subj').getValue()==txtNewAction){Ext.getCmp('subj').focus(true);alert(txtErrNoActionName);return false;}
	var newA=createActionOL(Ext.getCmp('subj').getValue());
	newA.Save();
	updateTheLatestThing();
	var id=newA.EntryID;
    Ext.getCmp("olid").setValue(id);
    cl=cl+1;
    //if (returnID==true){return id;}
    var gid=null;
    try{gid=g.getId();}catch(e){}
    if ((gid!="tgrid" && noGrid!=true) && (gid!="mgrid" && noGrid!=true))
    //called from grid
  	{var ar=new actionObject(newA);
  	var newRec=new actionRecord(ar);
  	var store=g.getStore();
    store.addSorted(newRec);
  	ids.push(newRec);cl++;
  	g.getSelectionModel().clearSelections();var nrr=new Array();nrr.push(newRec);g.getSelectionModel().selectRecords(nrr,true);
  	}

  	if (gid=="tgrid" && noGrid!=true)
  	//called from tickler grid
  	{var ar=new ticklerObject(newA);
  	var newRec=new ticklerRecord(ar);
  	var store=g.getStore();
    store.addSorted(newRec);
  	ids.push(newRec);cl++;
  	g.getSelectionModel().clearSelections();var nrr=new Array();nrr.push(newRec);g.getSelectionModel().selectRecords(nrr,true);
  	}

  	if (gid=="mgrid" && noGrid!=true)
  	//called from masterlist grid
  	{
  	var ar=new actionObject(newA);
  	ar.tag=txtNew;
    var newRec=new actionRecord(ar);
  	var store=g.getStore();
    store.insert(0,newRec);
  	ids.push(newRec);cl++;
  	g.getSelectionModel().clearSelections();var nrr=new Array();nrr.push(newRec);g.getSelectionModel().selectRecords(nrr,true);
  	}


	}

	for (var x=0;x<cl;x++)
	{
    if (noGrid!=true)
    {
      try{var id=ids[x].get("entryID");}catch(e){var id=olid;}
    }
    else
    {if (typeof(id)=="undefined"){var id=olid;}}


  if (notEmpty(id)){
	var it=NSpace.GetItemFromID(Trim(id));
  updateActionToOL(it);
	}

	//update store
    if (noGrid!=true)
     {
     var idx=gs.find("entryID",it.EntryID);
	     if (idx>-1){
       try{var ar=gs.getAt(idx);updateRecordItem(ar);}catch(e){}
       }
	 }
	 else
	 {
	 //not opened from grid. Check if there is a relevant grid to add the action
	 checkForExistingGridToUpdate(g,it);
	 }
	}

    try{var rval=Ext.getCmp("olid").getValue();}catch(e){}
    xmsg="<a class=jellolinktop onclick=scAction('"+it.EntryID+"')>"+it+"</a>";
    Ext.info.msg(txtActionEdit,xmsg+"<br>"+txtActionSuccess);
    if (returnID==true)
    {

    return rval;}
    else
    {
    Ext.getCmp("theactionform").destroy();
    }

updateTheLatestThing();

}

function updateActionToOL(it)
{
//update action to outlook task
    var oldNotes=Ext.getCmp("notes").oldvalue;
	var sv=Ext.getCmp('subj').getValue();
  if (sv!=txtMultiple){it.Subject=sv;}
	it.itemProperties.item(catProperty).Value=Ext.getCmp('tagfield').getValue();


	if (oldNotes!=Ext.getCmp('notes').getValue())
	{//if notes were changed update

		setJNotesProperty(it,Ext.getCmp('notes').getValue());

	}
	getJPriorityProperty(it);
	var upri=Ext.getCmp('pri').getValue();
	it.UserProperties.item("jpri").Value=upri;

  if (OLversion>11){
  //totalwork field//duration
  var durt=Ext.getCmp('totlwork').getValue();
  try{
  it.TotalWork=durt;}catch(e){}
  }

  if (upri>0){it.Importance=2;}
  if (upri==3){it.Importance=0;}
  if (upri==99){it.Importance=1;}
	var nv=Ext.getCmp('due').getValue();
	try{var dd=new Date (nv);var utc=dd.format('Y/m/d');
	if( it.Class == 43)
		it.itemProperties.item("TaskDueDate").Value = DisplayDate(dd);
	else
		it.itemProperties.item(dueProperty).Value=utc;
	}catch(e)
	{var utc="1/1/4501";
	if( it.Class == 43)
		it.itemProperties.item("TaskDueDate").Value = utc;
	else it.itemProperties.item(dueProperty).Value=utc;
	}

	it.Save();
}


function actionSelected(btn,vl, oneitem){
//j5
//apply an action to selected items
//if (btn.id!="adel"){return;}

status="Applying selected action to items...";
var cit;
if( typeof(oneitem) != "undefined" && oneitem != null)
	cit = oneitem;
else
	 cit=getCheckedItems();
if (cit==false){status=txtReady;return;}
var cl=cit.length;
if (cl==0){status=txtReady;return;}
if (btn.id=="adel" || btn.id=="acmdel"){if (choice=confirm(txtMsgDelItem)==false){status=txtReady;return;}}
var firstID=cit[0].get("entryID");
var firstIsTag=isidtag(firstID);
var editems=new Array();
var id="";var rmv=false;
var aFailed=false;
//for edit selected open form for bulk edit
if (btn.id=="aedit" || btn=="aedit" ||btn.id=="acmedit")
{
	if (cit[0].get("type")!="j"){editAction(cit);status=txtReady;return;}else{olItem(firstID);status=txtReady;return;}
}

	for (var x=0;x<cl;x++)
	{
	var id=cit[x].get("entryID");
	var itype=cit[x].get("type");
		 //delete
     if (btn.id=="adel" || btn.id=="acmdel")
      {
      if (itype=="a" || itype=="j"){OLDeleteItem(id);rmv=true;}
      if (itype=="t"){deleteTag(id,null,Ext.getCmp("grid"));rmv=true;}
      }
     //toggle NA
     if ((btn.id=="anext" || btn.id=="acmnext") && itype=="a"){aFailed=doAction(id);}
     //toggle Done
     if ((btn.id=="adone" || btn.id=="acmdone") && itype=="a"){aFailed=actionDone(id);rmv=true;}
     // move action to next step (INACTIVE)
	 if( btn.id == "astep")
		 	makeNextActionStep(id);
	 if(btn.id == "amakestep")
	 		editStep(id,cit[x]);
	 		//Convert to project
	 		if (btn.id=="aproject"){aFailed=convertToProject(id);}
	 		//copy action to another tag
	 		if (btn.id=="acopyaction"){editems.push(id);}
     //toggle Priority
     if (btn.id=="apri" && itype=="a"){aFailed=setActionPriority(id);}
     //add contacts
     if (btn.id=="adelg" && itype=="a"){editems.push(cit[x]);}
     //set priority by number key
     if (btn=="numaction" && itype=="a"){aFailed=setActionPriority(id,vl);}
     //toggle Start/stop
     if (btn.id=="astart" && itype=="a"){aFailed=startStopAction(id);}
     //toggle Privacy
     if (btn.id=="alock" && itype=="a"){aFailed=lockAction(id);}
     //open in outlook
     if (btn.id=="aoutlook" || btn.id=="acmoutlook"){olItem(id);}
     //create appointment
     if (btn.id=="aappoint"){appOfAction(id);}
     //create journal entry
     if (btn.id=="ajou"){addJournalInTag();}
     //tag item
     if (btn=="addtag")
     {
     var tg="";
     if (notEmpty(vl))
     {tg=Ext.getCmp("tagcmb"+vl).getRawValue();}
     else
     {tg=Ext.getCmp("tagcmb").getRawValue();}
     var it = NSpace.GetItemFromID(id);
     //addOLItemCategory(it,tg);
     if (notEmpty(tg)){if (addTagToAL(tg,it,cit[x],false)==false){status=txtReady;return;}}
     }


     //create shortcut
     if (btn.id=="ashort"){alShortcut(id,itype);}
     if( btn.id == "adue" || btn.id=="acmdue"){
	 	aFailed = true;
		try{var it = NSpace.GetItemFromID(id);
	 		if( typeof(it) != "undefined" && it != null)
	 		aFailed = setDueDate(it,btn.due);
		}catch(e){}
	}
     //edit selected
     if (btn.id=="aedit" || btn=="aedit" || btn.id=="acmedit"){editems.push(id);}

    if (btn.id!="ashort" && (btn.id!="aoutlook"  || btn.id=="acmoutlook") && btn.id!="adelg" && aFailed!=true)
      {
      //var g=getActiveGrid();
      updateRecordItem(cit[x],rmv);
      }

	}
if( btn.id == "acmemail" || btn.id == "aemail" && itype=="a") emailActionItem(cit);
if (btn.id=="adelg"){setActionContact(id,editems);}
if (btn.id=="acopyaction"){copyActionsToTag(id,editems);}
if (btn.id=="aproject"){refreshView();}

//if there is a form open, close it
try{var t=Ext.getCmp("theactionform");t.destroy();
//if action was not delete reopen form updated
if ((btn.id!="adel"  || btn.id=="acmdel") && (btn.id!="adone"  || btn.id=="acmdone")){editAction(cit);}else{Ext.getCmp("grid").show();}
}catch(e){}

status=txtReady;
}



function getCheckedItems(otherGrid,olview){
//j5
//return an array of checked grid items
var ar=new Array();
var gridid=null;
//grid selection
if (otherGrid==null)
{var g=getActiveGrid();}
else
{var g=Ext.getCmp(otherGrid);}

try{gridid=g.getId();}catch(e){}

if (gridid==null && typeof(olview)=="undefined"){olview="olnative";}

	if (notEmpty(olview))
	{
    if (olview.substr(0,4)=="olv:" || olview=="olnative")
    {
    	//outlook view selection
    	var sr=new Array();
    	try
    	{
    	 if (olview==null)
    	 {var olv=document.all.olnative.Selection;}
    	 else
    	 {
    	 var olm=document.getElementById(olview);
    	 try{var olv=olm.Selection;}catch(e){alert("IE 7 now blocks handling of Outlook view control contents. You can manage contents only by the built in keyboard shortcuts or context menus.");}
    	 }

    	 	for (var x=1;x<=(olv.Count);x++)
    		{
    		if (!notEmpty(olv.Item(x))){olv.Item(x).Subject=txtNoSubject;olv.Item(x).Save();}
    		sr.push(olv.Item(x));
    		}
		return sr;
    	}catch(e){Ext.info.msg(e+" "+txtAlOLCtlError,txtInvalid);return false;}
    }
	}

if (typeof(g)!="undefined")
{
//ext grid selection
var sr=g.getSelectionModel().getSelections();
var lastType=0;var flag=false;
	for (var x=0;x<sr.length;x++)
	{
	 var t=sr[x].get("type");
	 if (t=="t"){var ttype=1;}
	 else{ttype=2;}
	 if (lastType==0){lastType=ttype;}
	 if (ttype!=lastType){flag=true;break;}
	 lastType=ttype;
	}
	if (flag==true){alert(txtAlNotSameType);return false;}
}


return sr;
}

function quickAddAction(istag){
//j5
//quick action addition from toolbar
var qaf=Ext.getCmp('inlinenewaction');
var newAction=qaf.getValue();
if (newAction==null || newAction=="" || newAction==" " ){qaf.focus();return;}

var g=getActiveGrid();
if (trimLow(newAction.substr(0,4))=="ref:"){addJournalInTag(newAction.substr(4,200));qaf.reset();qaf.focus();return;}
if (trimLow(newAction.substr(0,1))=="#"){addJournalInTag(newAction.substr(1,200));qaf.reset();qaf.focus();return;}
if (istag==true){
//quick add an action
 var newA=createActionOL(newAction);
 newA.itemProperties.item(catProperty).Value=lastContext;
 newA.Save();
updateTheLatestThing();
 var it=newA;
 var ar=new actionObject(it);
 if (g.getId()=="mgrid"){ar.tag=txtUntagged;}
 var newRec=new actionRecord(ar);

/*
newAction += "\n";
if(jello.collectAdvancedMode==true || jello.collectAdvancedMode!=0)
	{
	importTextAdvanced(false,newAction,lastContext);
	}else{
	importText(false,newAction,lastContext);
	}
*/

}
else
{
//quick add a tag
jello.lastTagId++;
var tr=new tagRecord({
id:jello.lastTagId,
tag:newAction,
parent:lastOpenTagID,
istag:true,
notes:"",
isprivate:false
});
var ar=new actionObject(tr,true);
var newRec=new actionRecord(ar);
tagStore.add(tr);
tagStore.clearFilter();
syncStore(tagStore,"jello.tags");
//save settings
jese.saveCurrent();
var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();
	if (nd.firstChild!=null)
	{
	var nnd=new Ext.tree.TreeNode({text:newAction,id:jello.lastTagId,expandable:false,icon:imgPath+'list_components.gif',expanded:false});
	nd.appendChild(nnd);
	}
}

qaf.reset();
	var store=g.getStore();
	store.insert(0,newRec);

  g.getSelectionModel().selectFirstRow();
	try { updateAFooterCounter(store);}catch(e){}
	qaf.focus(true);

  updateTheLatestThing();
}

function createActionOL(newAction)
{  //j5
//create Action in outlook and return it
	var rits = NSpace.GetFolderFromID(jello.actionFolder).Items;
	var it=rits.Add();
	it.Subject=newAction;
	it.Status=jello.newActionStatus;
return it;
}

function NextActionToggleIcons(e,eid)
{
//j5
//render next action button
//var e=i.tags;
var vl="";
e=trimLow(e);
try{var ss=e.search(trimLow(getActionTagName()));}catch(b){var ss=-1;}
		if (ss != -1)
		{
		/*task is in action*/
		vl= "<a class=jellolink title='"+txtToggleNA+"' onclick=javascript:doAction(&quot;" + eid+ "&quot;,0)><img src="+imgPath+icIsNext+"></a>";
		}
		else
		{
		/*task NOT in action*/
		vl= "<a class=jellolink title='"+txtToggleNA+"' onclick=javascript:doAction(&quot;" + eid + "&quot;,1)><img src="+imgPath+icNoNext+"></a>";
		}

return vl;
}



function doAction(id,a)
{//j5
//set-unset task for next action (star)
try{var it=NSpace.GetItemFromID(id);}catch(e){alert(txtInvalid);return true;}
var cat=it.itemProperties.item(catProperty).Value;

var stat=0;
var cat2="";
try{var ss=trimLow(cat).search(trimLow(getActionTagName()));}catch(e){var ss=-1;}
if (a==null && ss>-1){stat=0;}
if (a==null && ss==-1){stat=1;}
if (a==0 && ss>-1){stat=0;}
if (a==0 && ss==-1){stat=1;}
if (a==1 && ss>-1){stat=0;}
if (a==1 && ss==-1){stat=1;}
    if (stat==0)
    {
    /*unset*/
    cat2=cat.replace(getActionTagName(),"");
    if (jello.nextToggleHighPri==1 || jello.nextToggleHighPri=="1"){setActionPriority(id,0);}
    }
    else
    {
    /*set*/
    cat2=cat + ";" + getActionTagName();cat2=cat2.replace(new RegExp(",","g"),";");
    if (jello.nextToggleHighPri==1 || jello.nextToggleHighPri=="1"){setActionPriority(id,1);}
    }

it.itemProperties.item(catProperty).Value=cat2;
it.Save();
	if (a!=null || typeof(a)!="undefined")
	{
	 //case of star push, update star
	 var st=document.activeElement;var rt="";
	 var g=getActiveGrid();
	var sr=g.getSelectionModel().getSelected();
	updateRecordItem(sr);

	}

}

function doFlag(id,a)
{//j5
//set-unset flag of item
if (a==null){a=1;}
try{
var it=NSpace.GetItemFromID(id);
it.FlagStatus=a;
it.Save();
//updateActionItem(id);
/*getPage();*/}catch(e){}
}

function clickActionCheckBox(id){
//j5
var t=document.getElementById(id);
markRow(t);
}


function actionDone(id)
{  //j5
//Toggle action as done

try{var pn=NSpace.GetItemfromID(id);}catch(e){alert(txtInvalid);return true;}
	if (pn.Complete==false){
		//move recurrent tasks to the next recurrence
		if (pn.isRecurring==true)
		{
		var skp=pn.SkipRecurrence();
		pn.Save();
		var t=new Date(pn.itemProperties.item(dueProperty).Value);
		var dd=DisplayDate(t);
		if (skp==true){Ext.info.msg(txtCompleted,"Task Recurrence completed. Task recreated for "+dd);return;}
		}
	
	pn.Complete=true;
    var nc="";
			e=pn.itemProperties.item(catProperty).Value;
			e=e.replace(new RegExp(",","g"),";");
			var ee=e.split(";");
			for (var x=0;x<ee.length;x++)
			{
				if (trimLow(ee[x])!=trimLow(getActionTagName())){
				nc=nc+ee[x]+";";}
			}
	pn.itemProperties.item(catProperty).Value=nc;
	}
	else{pn.Complete=false;}
	pn.Save();
}

function startStopAction(id)
{//j5
//start a task (in progress) or set it as not started
//for projects, set/unset inactive property
try{var it=NSpace.GetItemFromId(id);}catch(e){if (setTagArchived(id)==false){alert(txtInvalid);return true;}else{refreshView();return false;}}
if (it.Status==0){it.Status=1;}else{it.Status=0;}
it.Save();
return false;
}

function lockAction(id){
//j5
//toggle action private
try{var it=NSpace.GetItemFromId(id);}catch(e){if (setTagPrivacy(id)==false){alert(txtInvalid);return true;}else{refreshView();return false;}}
if (it.Sensitivity==0){it.Sensitivity=2;}else{it.Sensitivity=0;}
it.Save();
}

function actionObject(eID,other)
{//j5
//add action items (Tasks) to itms array

// if a Note
if (eID.Class==44)
{//notes need different handling
this.iclass=44;this.subject=eID.Subject;this.entryID=eID.EntryID;this.body=eID.body;this.created=new Date(eID.LastModificationTime);
this.icon=this.iclass+";0;0";
return;
}

if (eID.Class==42)
{
//journal entries
var jdd=new Date(eID.Start);
this.iclass=42;this.subject="("+DisplayDate(jdd)+") "+eID.Subject;this.entryID=eID.EntryID;this.created=new Date(eID.LastModificationTime);
this.icon=this.iclass+";0;0";
this.type="j";
try{this.tags=eID.itemProperties.Item(catProperty).Value;}catch(e){}
    if (jello.autoUpdateSecuredFields==false  || jello.autoUpdateSecuredFields==0 || jello.autoUpdateSecuredFields=="0")
    {
	this.body="<secured><a class=jellolinkTop onclick='getSecuredDataFromOutlook();'>"+txtTryToGetBody+"</a><br><a class=jellolinkTop onclick=getSecuredDataFromOutlook('"+this.entryID+"');>["+txtGetSecDataThisItm+"]</a></secured>";
    }else
    {
    this.body=eID.Body;
    }
try{this.attachmentCount=eID.Attachments.Count;}catch(e){}
return;
}

if (other!=null)
{
this.subject=eID.get("tag");
var itt=eID.get("istag");
var ipj=eID.get("isproject");
if (itt){this.iclass=100;}else{this.iclass=101;}
if (ipj){this.iclass=102;}
this.entryID=eID.get("id");
this.type="t";
	var ix=tagStore.find("id",new RegExp("^"+this.entryID+"$"));
	var r=tagStore.getAt(ix);
	if (r!=null)
	{this.body=r.get("notes");
  //this.notes=r.get("notes");
  }
return;
}
try{this.entryID=eID.EntryID;}catch(e)
  {//case of imap or other non downloaded items
    // also handles case of an item handled by an addon
  // that isn't a valid item
  try{this.subject=eID.Subject;}catch(e){this.subject=txtBadItem;};
  this.body=txtItOffline;
  this.created = this.due = new Date();
  this.iclass=43;
  this.icon="999;0;0";
  this.type="a";
  
  this.groupon = (eID.Class == 43? 1:0);
  
  return;
  }

this.unread=eID.UnRead;
this.subject=eID.Subject;
this.type="a";
this.groupon = (eID.Class == 43? 1:0);
this.contacts="";
    // class 43 == mailitem
  if (eID.Class==43)
  {

  try{this.markfordownload=eID.RemoteStatus;}catch(e){}
  if (eID.FlagStatus==2){this.flag=1;}
  if (eID.FlagStatus==1){this.flag=10;}

//this in case of security prompts of outlook
    if (jello.autoUpdateSecuredFields==false  || jello.autoUpdateSecuredFields==0 || jello.autoUpdateSecuredFields=="0")
    {
	this.body="<secured><a class=jellolinkTop onclick='getSecuredDataFromOutlook();'>"+txtTryToGetBody+"</a><br><a class=jellolinkTop onclick=getSecuredDataFromOutlook('"+this.entryID+"');>["+txtGetSecDataThisItm+"]</a></secured>";
    this.sender="?";
    }else
    {
    this.sender = eID.SenderName;
    this.to=eID.To;
	this.cc=eID.CC;
    }

	this.body = getItemBody(eID);
    this.notes = "";
    var ats = eID.Attachments;
    if( ats.Count > 0){
    	this.attachmentList = "";
    	var colon = "";
    	for( var i=1; i <= ats.Count; i++){
    		var z = ats.Item(i);
    		try{this.attachmentList += colon + "&nbsp;<a onclick=olItem('"+eID.EntryID+"'); class=jelloLink style='color:blue;text-decoration:underline;'>"+z.FileName+"</a>";}catch(e){}
    		colon=";";
    	}
	}

  }
  else
  {
      if (jello.autoUpdateSecuredFields==true || jello.autoUpdateSecuredFields==1 || jello.autoUpdateSecuredFields=="1")
      {
      this.body = getItemBody(eID);
      }
      else
      {
      this.body="<secured><a class=jellolinkTop onclick='getSecuredDataFromOutlook();'>"+txtTryToGetBody+"</a><br><a class=jellolinkTop onclick=getSecuredDataFromOutlook('"+this.entryID+"');>["+txtGetSecDataThisItm+"]</a></secured>";
      }
  this.notes=getJNotesProperty(eID);
  this.delState=eID.DelegationState;
  }
this.importance=getJPriorityProperty(eID);
if (this.subject==null || this.subject==""){this.subject=txtNoSubject;}
this.attachment=null;this.attachmentCount=0;
var mid=0;var did=0;
// if not a mail use create, else received
//if (eID.Class!=43){var t=new Date(eID.CreationTime);}else{var t=new Date(eID.ReceivedTime);}
var t;
// natural sort order for contact is fullname, but createtime should be used
var deftype=eID.parent.DefaultItemType;
if( deftype == 2)
{t = new Date(eID["CreationTime"]);}
else
{
if (deftype==0){t = new Date(eID.ReceivedTime);}
else{t = new Date(eID.CreationTime);}
}

this.created=t;
if (eID.Class==48){var tdc=new Date(eID.DateCompleted);}else{var tdc=null;}
this.Complete=tdc;
this.iclass=eID.Class;
this.sensitivity=eID.Sensitivity;
    try{var t=new Date(eID.itemProperties.item(dueProperty).Value);}catch(e){}
var dd=DisplayDate(t);if(dd==null || dd==""){dd="&nbsp;";}
this.due=t;
mid=(t.getMonth()+1);if (mid<10){mid="0"+mid;}
did=t.getDate();if (did<10){did="0"+did;}
this.status=eID.Status;
this.icon=this.iclass+";"+this.status+";"+this.sensitivity;
	// class 48 == task
	if (eID.Class!=48)
	//non task
	{this.due=new Date(getJDueProperty(eID));}
	else
	{//task
		
    try{
    if (eID.Links.Count>0){
		for (var x=1;x<=eID.Links.Count;x++)
		{this.contacts+=eID.Links.Item(x)+",";}
		if (notEmpty(this.contacts)){this.contacts=this.contacts.substr(0,this.contacts.length-1);}

		}}catch(e){}
		
			if (OLversion>11)
		  {
      //totalwork property for outlook 2007 or higher
      this.duration=eID.TotalWork;
      }
	}

	try{this.attachment=eID.UserProperties.Item("OLID").value;}catch(e){}
	try{this.attachmentCount=eID.Attachments.Count;}catch(e){}
	try{this.tags=eID.itemProperties.Item(catProperty).Value;}catch(e){}

}



function setActionPriority(id,num)
{
//set priority to an action
try{var it=NSpace.GetItemFromId(id);}catch(e){alert(txtInvalid);return true;}
var vl=getJPriorityProperty(it);
if(vl==null){vl=99;}
if (num==null){vl++;}else{vl=num;}
if (vl==4){vl=99;}
if (vl==100){vl=1;}
it.UserProperties.Item("jpri").Value=vl;

if (vl>0){it.Importance=2;}
if (vl==3){it.Importance=0;}
if (vl==99 || vl==0){it.Importance=1;}
it.Save();
}

function setActionContact(id,recs)
{
//set contacts to action
try{var it=NSpace.GetItemFromId(id);}catch(e){alert(txtInvalid);return;}
contactList=new Array();
	if (recs.length==1)
	{//if its one record to be updated, update contact list to show existing contacts
		var olid=recs[0].get("entryID");
		var it=NSpace.GetItemFromID(olid);
			
      try{
      if (it.Links.Count>0)
			{
				for (var x=1;x<=it.Links.Count;x++)
				{contactList.push(it.Links.Item(x).Item.EntryID);}
			}  }catch(e){}
	}
delegateItems(recs);
}


function addActionToolbar(istag,parentID,filtering,isProject){
//j5
//toolbar for action lists
var quicktext=txtQuickAction;
var upLevel=false;
if (!notEmpty(parentID)){parentID=0;}
upLevel=true;
if (istag==false){quicktext=txtQuickTag;}
var stepmenu = new Ext.menu.Menu({
        id: 'stepmenu',
        items: [{
        icon: 'img\\step.gif',
        cls:'x-btn-icon',
        id:'amakestep',
        text: txtCreateSteps,
			handler:actionSelected
      }]});

var menu2 = new Ext.menu.Menu({
        id: 'exportmenu',
        items: [{
                text: txtMnuSendPrw,
                tooltip:txtMnuSendPrwInfo,
                icon: 'img\\mail.gif',
                handler:function(){
        printAList(false);}
        },{
                text: txtPrintDirectly,
                tooltip:txtPrintDirectlyInfo,
                icon: 'img\\print.gif',
                handler:function(){
        printAList(true);}
        },{
        text: txtPlainTextExp,
        tooltip:txtPlainTextExpInfo,
        icon: 'img\\note.gif',
        handler:function(){printToText();}
        },{
                text: txtFExport,
                tooltip:txtFExportInfo,
                icon: 'img\\move.gif',
                handler:function(){

           Ext.Msg.show({
           title:txtFExport,
           msg: txtFExportInfo + '<br>' + txtFExportSel,
           buttons: {yes:txtAllItems,no:txtSelItems,cancel:txtCancel},
           fn: function(b,t){exportToFolder(b);},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });

        }
        },{
        text: txtPtSndWTicklers,
        tooltip:txtPtSndWTicklers,
        id:'calprt',
        icon: 'img\\calendar.gif',
        handler:function(){printWithTicklersForm();}
        },{
        text: txtPtProjHistory,
        tooltip:txtPtProjHistory,
        id:'prjprt',
        hidden:!isProject,
        icon: 'img\\page_bookmark.gif',
        handler:function(){printProjectHistory();}
        }


        ]}
    );

 var menu3 = new Ext.menu.Menu({
        id: 'actionsmenu',
        items: [{
                text: txtItmAddToContact,
                cls:'x-btn-icon',
                id:'adelg',
                icon: 'img\\user.gif',
                handler: actionSelected
        },{
                text: txtsetPriority,
                cls:'x-btn-icon',
                id:'apri',
                icon: 'img\\priority.gif',
                handler: actionSelected
        },{
                text: txtStartStop,
                cls:'x-btn-icon',
                id:'astart',
                icon: 'img\\widget-add.gif',
                handler:actionSelected
        },{
                text: txtSetPrivacy,
                cls:'x-btn-icon',
                id:'alock',
                icon: 'img\\page_lock.gif',
                handler:actionSelected
        },
        {
                text: txtTaskCnvToPrj,
                cls:'x-btn-icon',
                id:'aproject',
                icon: 'img\\project.gif',
                handler:actionSelected
        },{
                text: txtCopy+' to another Tag',
                cls:'x-btn-icon',
                id:'acopyaction',
                icon: 'img\\copy.gif',
                handler:actionSelected
        },{
        icon: 'img\\forward.gif',
        cls:'x-btn-icon',
        id: 'aemail',
        text: txtStatusReport,
        visible:OLversion>10,
        handler : actionSelected
        },{
        icon: 'img\\appoint.gif',
        cls:'x-btn-icon',
        text: txtItmAppInfo,
        id:'aappoint',
        handler : actionSelected
        },{
        icon: 'img\\shcut.gif',
        cls:'x-btn-icon',
        id:'ashort',
        text: txtShcutInfo,
        handler : actionSelected
        },{
        icon: 'img\\journal.gif',
        cls:'x-btn-icon',
        id:'ajou',
        text: txtJournalInfo,
        handler : actionSelected
        }
        ]});

    tbar1 = new Ext.Toolbar();
    //tbar1.render('toolbar');
        tbar1.add({
        icon: 'img\\new.gif',
        cls:'x-btn-icon',
        hidden:!istag,
        tooltip: txtAddNewItmInfo,
			listeners:{
            click: function(b,e){
			e.stopEvent();
			e.preventDefault();
			e.stopPropagation();
			editAction(null);
            }
            }
        },
        {
        icon: 'img\\folder_new.gif',
        cls:'x-btn-icon',
        id:'tgnew',
        tooltip: txtAddNewTagInfo,
			listeners:{
            click: function(b,e){
			e.stopEvent();
			e.preventDefault();
			e.stopPropagation();
			editTag(null,istag);
            }
            }
        }
        ,'-',{
        icon: 'img\\notes.gif',
        cls:'x-btn-icon',
        tooltip: 'Toggle in-line editing of action subjects',
        id:'aeditable',
        enableToggle:true,
        pressed:false,
        handler : toggleEditableGrid
        },{
        icon: 'img\\page_edit.gif',
        cls:'x-btn-icon',
        tooltip: txtEdSelItmInf,
        id:'aedit',
        handler : actionSelected
        },
        {
        icon: 'img\\page_delete.gif',
        cls:'x-btn-icon',
        tooltip: txtDelItmInfo,
        id:'adel',
        handler : actionSelected
        },'-',{
        icon: 'img\\IsNext.gif',
        cls:'x-btn-icon',
        id:'anext',
        tooltip: txtSetNAInfo,
        handler : actionSelected
        },{
        icon: 'img\\check.gif',
        cls:'x-btn-icon',
        id:'adone',
        tooltip: txtCompleteItmInfo,
        handler : actionSelected
        },
        '-',

        {
            text:'Actions',
            menu: menu3
        },

        {
                tooltip: txtOpenInOutlook + '(Ctrl+Q)',
                cls:'x-btn-icon',
                id:'aoutlook',
                icon: 'img\\info.gif',
                handler:actionSelected
        },'-',
        {
    xtype: 'textfield',
		width:170,
		emptyText:quicktext,
		id:'inlinenewaction',
		listeners:{
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                quickAddAction(istag);}},
        focus: function(t){t.getEl().fadeIn();}
		}},'-',{
        icon: 'img\\link.gif',
        cls:'x-btn-icon',
        tooltip: 'Up level to Parent',
        hidden:!upLevel,
        id:'auplevel',
        listeners:{click:function(){tagUpLevel(parentID);}}
        },
		/*{
        icon: 'img\\filter.gif',
        cls:'x-btn-icon',
        tooltip: 'Set A View',
        hidden:!upLevel,
        id:'setvw',
        listeners:{click:function(){setActionView();}}
        },*/

        {
            text:txtMenuOtherFunctions,
            menu: menu2
        },'-','->',{
        icon: 'img\\refresh.gif',
        cls:'x-btn-icon',
        tooltip: txtRefreshInfo+' (Ctrl+S)',
        id:'arefresh',
        handler : function(){refreshView();
        }
        }

         );

         	 var fltCmb=new Ext.form.ComboBox({
                id: 'filter',
                store:filterValues,
                hideTrigger:false,
                valueField:'value',
                displayField:'text',
                width:80,
                style:'font-size:9px;',
                listWidth:220,
                triggerAction:'all',
                selectOnFocus:true,
				editable:false,
                emptyText:txtFilter+'...',
                mode:'local',
                hidden:!istag,
                value:filtering,
                listeners:{
                specialkey: function(f, e){
				if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                filterActions();}},
				select: function(cb,rec,idx){
	            filterActions();
		        }
                }
                });
tbar1.insert(18,fltCmb);


var tagCmb=new Ext.form.ComboBox({
                //fieldLabel: 'Tags',
                id: 'tagcmb',
                store:globalTags,
                hideTrigger:false,
                typeAhead:false,
                triggerAction:'all',
                width:80,
                listWidth:220,
                emptyText:txtTagSelInfo,
                mode:'local',
                maxHeight:250,
                listeners:{
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                actionSelected("addtag");
                Ext.getCmp("tagcmb").reset();

                }},
             select: function(cb,rec,idx){
            actionSelected("addtag");
            Ext.getCmp("tagcmb").reset();
            setTimeout(function(){
            Ext.getCmp("tagcmb").focus();
            },300);
            }
		}

            });
tbar1.insert(9,tagCmb);
      return tbar1;
}


function getIcon(c,m,r){
//j5
//assign icon to an outlook item
var ic=icTask;
var stl="";
try{
var t=c.split(";");
c=t[0];
var cst=t[1];
var cse=t[2];
if (c==43)
	{
	ic=icMail;var icem="<img src=img\\"+ic+">";
	//follow up flag icons
	var fl=r.get("flag");
	if (fl==1){icem="<span style='background-image:url("+imgPath+ic+");background-repeat:no-repeat;filter:alpha(OPACITY=20)'><img title='Flagged' src=img\\flag.gif></span>";}
	if (fl==10){icem="<span style='background-image:url("+imgPath+ic+");background-repeat:no-repeat;filter:alpha(OPACITY=20)'><img title='Flagged' src=img\\check.gif></span>";}
	return icem;
	}

if (c==44){ic=icJournal;}
if (c==null){ic=icApp;return "<img src=img\\"+ic+">";}
if (c==48){
if (cst==0 && jello.showNotStartedAlways==false){ic=icTaskNS;}
if (cse==2){ic=icTaskP;}}
if (c==999)
{//encrypted or unable to read items
ic="icon_key.gif"
}
ic=iconBasedOnClass(c,ic);
//attached items (for multi action tasks!) Try to have this in action object
var sic=icTask;
var att=r.get("attachment");
if (att!=null)
{
  try{
  var olit=getFirstAttached(att);
    if (notEmpty(olit))
    {
    var ci=olit.Class;sic=iconBasedOnClass(ci);
    stl="<span style='background-image:url("+imgPath+sic+");background-repeat:no-repeat;filter:alpha(OPACITY=20)'>&nbsp;&nbsp;";
    }
    else{sic="";stl="";}
    }catch(e)
    {var olit=null;}

}

}catch(e){}
if (typeof(t)=="undefined")
{//tag icons
var i=r.get("iclass");
var eeid=r.get("entryID");
var prv=getTagPrivacy(eeid);
var arv=getTagArchived(eeid);
var ctic=getTagIcon(eeid);
if (i==101){ic="folder.gif";if(prv){ic="folder_lock.gif";}}
if (i==100){ic="list_components.gif";if(prv){ic="page_lock.gif";}}
if (i==102){ic="project.gif";if(arv){ic="box.gif";}}
if (prv){ic="page_lock.gif";}
if (notEmpty(ctic)){ic=ctic;}
}
try{if (isOLReccuring(r.get("entryID"))){ic="taskrec.gif";}}catch(e){}
rt=stl+"<img src='img\\"+ic+"'>";
if (notEmpty(stl)){rt+="</span>";}
return rt;
}

function actionListTitleIcon(ctx){
//j5
//return context or project icon
try{var ix=tagStore.find("tag",ctx);
var r=tagStore.getAt(ix);
var icon=r.get("icon");
if (notEmpty(icon)){return "<img src='img\\"+icon+"' style=float:left;>";}
var isTag=r.get("istag");
var isPrj=r.get("isproject");
var prv=r.get("private");
var arc=r.get("archived");
var id=r.get("id");
var icc="list_components.gif";
if (!isTag){icc="ffolder.gif";}
if (prv){icc="page_lock.gif";}
if (isPrj){icc="project.gif";}
if (arc){icc="box.gif";}
if (prv && !isTag){icc="folder_lock.gif";}
switch (id)
{
case 6:icc="isnext.gif";break;
case 7:icc="list_review.gif";break;
case 8:icc="list_waiting.gif";break;
case 9:icc="list_someday.gif";break;
}
return "<img src=img\\"+icc+" style=float:left;>";
}catch(e){return "<img src=img\\icon_extension.gif style=float:left;>";}
}

function OriggetJNotesProperty(it){
//get custom notes field value. Create if not exists
var jn="";
try{
jn=it.UserProperties.Item("jnotes").Value;
}catch(e){it.UserProperties.Add("jnotes",1,true);}

return jn;
}

function getJNotesProperty(it){
 /* FaultyJusers */
//this pops up security prompts at all times
//get custom notes field value. Create if not exists
var jn="";
if (OLversion >= 12 && jello.autoUpdateTaskNotes==true && jello.autoUpdateSecuredFields==true)
{
	// notes are in the body
	var oln = getOLBodySection(it,jelloNotesDelim);
	if(oln != null)
		jn = oln;
	else{
		// didn't find notes in body, see if in custom field
		var itm = it.UserProperties.Find("jnotes",true);
		if( itm != null )
			jn = itm.Value;
	}
}else{
try{
jn=it.UserProperties.Item("jnotes").Value;
}catch(e){it.UserProperties.Add("jnotes",1,true);}
}
return jn;
}
function getItemBody(it)
{
	if( it.Class == 43){
		var oln="";
		//this in case of security prompts of outlook
	    if (jello.autoUpdateSecuredFields==false  || jello.autoUpdateSecuredFields==0 || jello.autoUpdateSecuredFields=="0")
	    {oln="<secured><a class=jellolinkTop onclick='getSecuredDataFromOutlook();'>"+txtTryToGetBody+"</a><br><a class=jellolinkTop onclick=getSecuredDataFromOutlook('"+it.EntryID+"');>["+txtGetSecDataThisItm+"]</a></secured>";
	    }else if (jello.autoUpdateSecuredFields==true || jello.autoUpdateSecuredFields==1 || jello.autoUpdateSecuredFields=="1")
	    {
	    //otherwise this
	    if (jello.mailPreviewHTML==true){oln=it.HTMLBody;}else{oln=cleanBody(it.Body);}
    }
	return oln;


	}else{

if (jello.autoUpdateSecuredFields==true || jello.autoUpdateSecuredFields==1 || jello.autoUpdateSecuredFields=="1")
{
		if (jello.autoUpdateTaskNotes==false)
			return cleanBody(it.Body);
		var oln=cleanJelloSectionFromOLBody(it.body,jelloNotesDelim);
		return cleanBody(oln);
	}else
  {
  var oln="<secured><a class=jellolinkTop onclick='getSecuredDataFromOutlook();'>"+txtTryToGetBody+"</a><br><a class=jellolinkTop onclick=getSecuredDataFromOutlook('"+it.EntryID+"');>["+txtGetSecDataThisItm+"]</a></secured>";
  return getJNotesProperty(it)+"<br>"+oln;
  }

  }
	return "";

}
function setJNotesProperty(it,val)
{
if (OLversion >= 12 && jello.autoUpdateTaskNotes==true)
{
	setOLBodySection(it,jelloNotesDelim,val);

}else{
	var up = it.UserProperties.Find("jnotes");
	if( up == null)
		up = it.UserProperties.Add("jnotes",1,true);
	up.Value=val;
}
}
function buildGrid(ctx,iits,counter,filtering,oItems,istag,storeOnly)
{
//grid for action records
// ctx == null asks for a popup
status=txtStatusGetting="...";
thisGrid="grid";
if( ctx != null)initScreen(true,"showContext('"+ctx+"',1)");
var gridName = (ctx == null?"actgrid":"grid");
var isWaiting=false;
var listStore;
var assignedView="";
var isProject = false;

if( ctx != null){
    var ttt=getTagID(ctx);
    isProject=getTagProject(ttt);
    if(ttt===8)
    {
    listStore=buildWaitingGrid(ctx,iits,counter,filtering,oItems,istag);
    isWaiting=true;
    assignedView=new Ext.grid.GroupingView({forceFit:true,groupTextTpl: '{text} ({[values.rs.length]} {[values.rs.length > 1 ? txtItems : txtItem]})'});
    }
}
getTagsArray();
var g=getActiveGrid();
//try{g.destroy()}catch(e){}

if (!isWaiting)
{
	listStore = new Ext.data.SimpleStore({
        id:'actStore',
        fields: [
           {name: 'subject'},
           {name: 'entryID'},
          {name: 'type'},
		   {name: 'importance'},
		   {name: 'attachment'},
		   {name: 'created',type:'date'},
		   {name: 'iclass'},
		   {name: 'sensitivity'},
		   {name: 'due',type:'date'},
		   {name: 'status'},
		   {name: 'icon'},
		   {name: 'body'},
		   {name: 'unread',type:'boolean'},
		   {name: 'tags'}
        ]
    });


	//add other type items
	if (oItems!=null)
	{
		for (var x=0;x<oItems.length;x++)
		{
			var it=null;
			try{it=NSpace.GetItemFromID(oItems[x].get("entryID"));}catch(e){}

			var ar=null;
			if (!notEmpty(it)){ar=new actionObject(oItems[x],true);}
			else{ar=new actionObject(it);}
			var newRec=new actionRecord(ar);
			listStore.add(newRec);
		}
	}

	//add actions
	if (iits!=null)
	{
		if (iits.Count>0)
		{
			for (var x=1;x<=iits.Count;x++)
			{
				var ar=new actionObject(iits(x));
				var newRec=new actionRecord(ar);
				listStore.add(newRec);
			}
		 listStore.sort("importance","ASC");
		 }

	}
}

if (storeOnly){return listStore;}

var gridwaitMask = new Ext.LoadMask(Ext.getBody(), {msg:txtMsgWait},{store:listStore});
sm= new Ext.grid.CheckboxSelectionModel({});

//insert reference items
if( ctx != null) addReferenceItems(ctx,listStore);

 var cc=listStore.getCount();

if( ctx != null)updateCounter(cc);
if (filtering!="0"){updateCounter("("+txtFilterOn+") "+cc);}
var fltr=ctx;
var dts=jello.dateSeperator;
var ALinline=false;if(jello.ALinlineduedate==1 || jello.ALinlineduedate=="1"){ALinline=true;}
var fmt="n"+dts+"j"+dts+"y";
//context
if (filtering=="99"){fltr=lastContext+" "+txtFilterOn+" "+ ctx;}
var hasParent=null;
if( ctx != null) gasParent = getTagParent(ctx);
var gridTBar=null;
if( ctx != null) gridTBar=addActionToolbar(istag,hasParent,filtering,isProject);
var durfield={header:'Total Work (unavailable)'};

if (OLversion>11){
durfield={header: 'Total Work', width: 100,hidden:true, sortable: true, renderer:DisplayDuration, dataIndex: 'duration'};
}

    var grid = new Ext.grid.EditorGridPanel({
        store: listStore,
        id:(ctx == null?'actgrid':'grid'),
        tbar:gridTBar,
        clicksToEdit: 1,
        columns: [
        sm,
			{header: "Priority", width: 35, fixed:true, sortable: true, renderer: getImportance, dataIndex: 'importance'},
			{header: "", width: 30, hideable:false, fixed:true, sortable: false, renderer: getIcon, dataIndex: 'icon'},
            {header: "Sensitivity", width: 20, hidden:true, fixed:false, sortable: true, dataIndex: 'sensitivity'},
            {header: txtCreated, width: 80, hidden:true, sortable: true, renderer:DisplayAppDate, dataIndex: 'created'},
            {header: txtSubject, width: 380, sortable: true, renderer: renderSubject,dataIndex: 'subject',
            editable:false,editor: new Ext.form.TextField({
            id:'inlineaction',
            style:{'font-size': '12px','font-weight':'bold'}
            })
            },
            //{header: txtDueDate, width: 80, sortable: true,hidden:!istag, renderer:DisplayAppDate, dataIndex: 'due'},
            {header: txtDueDate, width: 80, sortable: true,hidden:!istag, renderer:DisplayAppDate, dataIndex: 'due',
                editable: ALinline,editor: new Ext.form.DateField({
                    format: fmt,
                    id:'inlinedate',
                    editable: false
                })
             },
            {header: txtTags, width: 100,hidden:true, sortable: true, dataIndex: 'tags'},
			{header: txtContacts, width: 100,hidden:true, sortable: true, dataIndex: 'contacts'},
			durfield,
			{header: txtCompleted, width: 80,hidden:true, sortable: true, renderer:DisplayAppDate, dataIndex: 'completed'},
            {header: txtNotes, width: 100,hidden:true, sortable: true, dataIndex: 'body'}
        ],
        stripeRows: true,
        //columnLines:true,
        autoScroll:true,
        enableColumnHide:true,
        split: true,
        deferRowRender:false,
        viewConfig:{
        emptyText:txtGridNoData
        },
        region: 'north',
        trackMouseOver:true,
        //autoHeight:true,
        //height:Ext.getBody().getViewSize().height-jello.actionPreviewHeight,
        height:200,
        view:assignedView,
		sm:sm,
    listeners:{
    mouseover: function(e){thisGrid='grid';},
		rowdblclick: function(g,row,e){
		openGridRecord(g);
        },
			cellcontextmenu: function(g, row,cell,e){

 			e.stopEvent();
			e.preventDefault();
			e.stopPropagation();
			g.getSelectionModel().selectRow(row);
			rightClickItemMenu(e,row,g);
        },
        afteredit: function (obj){
                actionDateEdited(obj);
        }
		},
        enableColumnMove:true,
        border:true,
        loadMask:gridwaitMask,
        layout:'fit',
        ctx: ctx

    });

    
  //render to portal
    if( ctx != null){
    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.setAutoScroll(false);
    ppnl.doLayout();
    }else{
    	var win = new Ext.Window({
        title: "Actions",
        width: 520,
        height:370,
        id:'actpop',
        minWidth: 300,
        minHeight: 280,
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: grid,
        modal:true,
        listeners: {destroy:function(t){
			// if a tag view, refresh when we exit
			var g = Ext.getCmp("grid");
			if(  typeof(g) != "undefined" && g != null)
				setTimeout(function(){refreshView();},50);
			showOLViewCtl(true);}}

        });

        showOLViewCtl(false);
        win.show();
        win.setActive(true);

    }


    	grid.getSelectionModel().on('rowselect', function(sm, rowIdx, r) {
		var detailPanel = Ext.getCmp('previewpanel');
		actionTpl.overwrite(detailPanel.body, r.data);
				var gb=r.get("body");
	   });


grid.on('columnresize',function(index,size){saveGridState("grid");});
grid.on('sortchange',function(){saveGridState("grid");});
grid.getColumnModel().on('columnmoved',function(){saveGridState("grid");});
grid.getColumnModel().on('hiddenchange',function(index,size){saveGridState("grid");});
try{restoreGridState("grid");}catch(e){}

if (isWaiting)
{
listStore.groupBy('contacts',true);
}

    //grid.render('list');



	 if (filtering=="99"){

		var cspan=document.getElementById("jelloCounter");
		cspan.innerHTML+="&nbsp;<a onclick=javascript:showContext(lastContext,1);>"+txtRmvTgFilter+"</a>";

		 }

setTimeout(function(){
try{
if (jello.selectFirstItem==1 || jello.selectFirstItem=="1"){grid.getSelectionModel().selectRow(0);}
grid.getView().focusRow(0);grid.focus();}catch(e){}
resizeGrids();
},100);

// map one key by key code
var lmap1 = new Ext.KeyMap(gridName, {
    key: 13,
    fn: function(){openGridRecord(Ext.getCmp("grid"));},
    scope: this
});

//insert focuses the new action toolbar box
var lmap2 = new Ext.KeyMap(gridName, {
    key: Ext.EventObject.INSERT,
    fn: function(){Ext.getCmp("inlinenewaction").focus();try{var aw=Ext.getCmp("subj");aw.focus();}catch(e){}},
    stopEvent:true,
    ctrl:false,
    shift:false,
    scope: this
});
//ctrl + INS adds action by form
var lmap3 = new Ext.KeyMap(gridName, {
    key: Ext.EventObject.INSERT,
    fn: function(){editAction(null);},
    ctrl:true,
    shift:false,
    stopEvent:true,
    scope: this
});
//DEL deletes items
var lmap4 = new Ext.KeyMap(gridName, {
    key: Ext.EventObject.DELETE,
    fn: function(){
		if (document.activeElement.id!="inlinenewaction"){
		actionSelected(Ext.getCmp("adel"));
		}
    },
    stopEvent:false,
    ctrl:false,
    scope: this
});

//Ctrl+q open OL item
var lmap13 = new Ext.KeyMap(gridName, {
    key: 'q',
    fn: function(){
    actionSelected(Ext.getCmp("aoutlook"));
    },
    stopEvent:true,
    ctrl:true,
    scope: this
});

//ctrl + T edits active tag
var lmap5 = new Ext.KeyMap(gridName, {
    key: 't',
    fn: function(){
    var r=getTag(ctx);var tid=r.get("id");
    editTag(tid);},
    ctrl:true,
    stopEvent:true,
    scope: this
});

//Ctrl+E assigns contacts
var lmap6 = new Ext.KeyMap(gridName, {
    key: "e",
    ctrl:true,
    fn: function(){
	actionSelected(Ext.getCmp('adelg'));
   },
    stopEvent:true,
    scope: this});

//Ctrl+J/K moves to previous/next item
var lmap7 = new Ext.KeyMap(gridName, {
    key: 'k',
    fn: function(){gridGo(1);},
    stopEvent:true,
    ctrl:true,
    scope: this
});
var lmap8 = new Ext.KeyMap(gridName, {
    key: 'j',
    ctrl:true,
    fn: function(k,e){
	gridGo(-1);
    },
    stopEvent:true,
    ctrl:true,
    scope: this
});
var lmap13 = new Ext.KeyMap(gridName, {
    key: 's',
    fn: function(){refreshView();},
    stopEvent:true,
    ctrl:true,
    scope: this
});

//123 and 0 keys are setting priority
if (jello.enablekeys03==1 || jello.enablekeys03=="1")
  {
  var lmap9 = new Ext.KeyMap(document, {key: '1',fn: function(){actionSelected("numaction",1);},ctrl:true,scope: this});
  var lmap10 = new Ext.KeyMap(document, {key: '2',fn: function(){actionSelected("numaction",2);},ctrl:true,scope: this});
  var lmap11 = new Ext.KeyMap(document, {key: '3',fn: function(){actionSelected("numaction",3);},ctrl:true,scope: this});
  var lmap12 = new Ext.KeyMap(document, {key: '0',fn: function(){actionSelected("numaction",99);},ctrl:true,scope: this});
  }

if( ctx != null) ppnl.setTitle(actionListTitleIcon(ctx)+" "+TagFullPath(hasParent,ctx)+" "+tagDescription(ctx));
status=txtReady;
}

function actionContextMenu(rec, id, inlist)
{
try{Ext.getCmp("acmrclickmenu").destroy();}catch(e){}
	var nmnu= new Ext.menu.Menu({
		id: 'acmrclickmenu',
		items: [
    {
        icon: 'img\\page_edit.gif',
        text: txtEdSelItmInf,
        id:'acmedit',
        listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
        },'-',
		{
        icon: 'img\\page_delete.gif',
        id:'acmdel',
        text: txtDelItmInfo,
        listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
        },
		{
		text: txtSetNAInfo,
		icon: 'img\\IsNext.gif',
		id:'acmnext',
		listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
		},
		{
        icon: 'img\\check.gif',
        id:'acmdone',
        text: txtCompleteItmInfo,
        listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
        },'-',{
		text: txtBody,
		id:'acmbody',
		handler:function(){
		popTaskBody(id,rec);}
		},
    {
  text: txtStatusReport,
  visible:OLversion>10,
  id:'acmemail',
  icon: 'img\\forward.gif',
  listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
  },
		{
            icon: 'img\\calendar.gif',
            text:txtDueForInfo,
            id:'acmpopdate',
            menu: new Ext.menu.DateMenu({
        	handler : function(dp, idate){
            //var dt=idate.format('m/j/Y');
            //if(jello.dateFormat == 0 || jello.dateFormat == "0"){dt=idate.format('j/m/Y');}
            var dt=idate.toUTCString();
            var b = new Ext.Button({id: 'acmdue'});
            b.due = dt;
            actionClickForwarder(b,null,inlist,rec);
        	}
    		})
        },
		{
        text: txtOpenInOutlook,
        id:'acmoutlook',
//        hidden: (getActiveGrid().getId() == "actgrid"?true:false),
        icon: 'img\\info.gif',
        listeners:{click: function(b,e){actionClickForwarder(b,e,inlist,rec);} }
        }
        ,{
			text: txtShowWithTag,
			id: 'avsametasks',
			icon: 'img\\task.gif',
			hidden: (getActiveGrid().getId() == "actgrid"?true:false),
			listeners: {click: function(b,e){tasksWithTag(rec);} }
		}

		]
	});

	return nmnu;
}
function actionClickForwarder(b,e,inlist, rec)
{
	if( inlist){
			actionSelected(b,e);
	}else{
		var x = new Array();
		x.push(rec);
		actionSelected(b,null,x);
	}
}

function getImportance(v,m,r)
{
//get item's importance and return color, and attachment icon next to it
  var acou=r.get("attachmentCount");
  var aticon="";
  var id = r.get("entryID");
  if (acou>0){aticon="<a class=jellolink title='"+acou+" "+txtAttachments+"' onclick='openInboxItem(\""+id+"\");'><img src="+imgPath+"attach.gif></a>";}
  var ic="";
  if (v==1){ic="red";}
  if (v==2){ic="blue";}
  if (v==3){ic="gainsboro";}
  return "<span style=background:"+ic+">&nbsp;</span>"+aticon;
}

function renderSubject(v,m,r)
{
//render the subject line in grid
var tg=r.get("tags");
var id=r.get("entryID");
var st=r.get("status");
var cl=r.get("iclass");
var un=r.get("unread");
var con=r.get("contacts");
var dstate=r.get("delState");
  if (cl>99){
  var cnt="";
  try{cnt="&nbsp;("+showContext(v,0)+")";}catch(e){}
  return "<b>"+Ext.util.Format.htmlEncode(v)+"</b>"+cnt;
  };
var rt="";
var tList=tagList("formtagdisplay",lastContext,tg,id);

  var subj=Ext.util.Format.htmlEncode(v);
  if (st==2){subj="<strike>"+subj+"</strike>";}
  if (un==true){subj="<b>"+subj+"</b>";}
  if (tList[1]==true){subj="<b>"+subj+"</b>";}
  var delreply="";
  if (dstate==1){delreply=txtDelgStatSent;}
  if (dstate==2){delreply=txtDelgStatAcc;}
  if (dstate==3){delreply=txtDelgStatDec;}
  if (dstate>0){subj="<img src=img//inboxmail.gif><font size=1 color=green>("+delreply+")</font> "+subj;}

  if (cl==42){rt=subj+"</span>&nbsp;"+tList[0];return rt;}
  var conicon="";
  var con2=con.replace(new RegExp("'","g")," ");
  if (notEmpty(con) && con!="("+txtContactNo+")"){conicon="<a class=jellolink title='"+con2+"'><img src="+imgPath+"user.gif></a>";}

  rt+=NextActionToggleIcons(tg,id)+conicon+"&nbsp;"+subj+"&nbsp;"+tList[0];

 return rt;
}

function filterActions(){
//filter actions list
var f=Ext.getCmp("filter").getValue();
var thisG=getActiveGrid();
if (thisG.id=="mgrid")
{pMaster(f);}
else
{
//update filter to be available next time
var tid=getTagID(lastContext);
	if (tid!=0)
	{
	setTagFilter(tid,f);
	}
//show tag with filter
showContext(lastContext,1,null,f);
}

}

function openGridRecord(g){
//open row record for editing
		try{
		var sr=g.getSelectionModel().getSelections();
			if (sr.length==1)
			{//if selection is one tag navigate instaed of showing props
			 var tp=sr[0].get("type");
			 var id=sr[0].get("entryID");
				if (tp=="t")
				{
					 var ix=tagStore.find("id",new RegExp("^"+id+"$"));
					 var r=tagStore.getAt(ix);
setTimeout(function(){
           showTagFolder(r);
           },6);

			 return;
			}
      if (tp=="a"){actionSelected("aedit");}
      if (tp=="j"){olItem(id);}
		}}catch(e){}


}


function updateAFooterCounter(s)
{
  updateCounter(s.getCount());
}

function refreshView()
{
var g=getActiveGrid();

    if (g.getId()=="mgrid")
    {
    pMaster();
    return;
    }

    if (g.getId()=="ggrid")
    {
    tagManager();
    return;
    }

    if (g.getId()=="grid")
    {
	if(lastContext=="##Inbox"){showNoMailInbox(getInboxTaskDASL());return;}
    var tt=getTagType(lastOpenTagID);
    if (tt)
    {
    showContext(lastOpenTagID,1);
    }
    else
    {showTagFolder(null,lastOpenTagID);}
    }


}

function getActionAttachmentLink(it)
{
//return a link to an action's attachment with icon
	var ret="";
	try{
	var aid=it.UserProperties.Item("OLID").value;
	var aids=aid.split(";");
	 for (var x=0;x<aids.length;x++)
	 {
	 var ait=NSpace.GetItemFromID(aids[x]);
	 var ic="";var c=ait.Class;
			if (c==26){ic=icApp;}
			if (c==40){ic=icContact;}
			if (c==43){ic=icMail;}
			if (c==44){ic=icNote;}
			if (c==42){ic=icJournal;}
			if (c==0){ic="mng.gif";}
var attName=ait.Subject;
if (!notEmpty(attName)){attName=txtNoSubject;}
ret+="<a style='background-image:url("+imgPath+ic+");background-repeat:no-repeat;padding-left:18px;width:200;height:25;color:blue;' class=jellolink onclick=olItem('"+aids[x]+"') title='Parent:"+ait.Parent.FolderPath+" ["+txtClickOpen+"]'><b>"+attName+"</b></a>&nbsp;";
ret+="<a style='background-image:url("+imgPath+"page_bookmark.gif);background-repeat:no-repeat;padding-left:18px;width:200;height:25;color:blue;' class=jellolink onclick=OLFolderOpenNewWindow('"+ait.Parent.EntryID+"'); title='Click to open folder'>"+ait.Parent.FolderPath+"</a><br>";
    }
		}catch(e)
		{}

	 //return contacts linked with this action
 try{
 var lks=it.Links.Count;
  if (lks>0)
  {
		for (var x=1;x<=lks;x++)
		{
		var lit=it.Links.Item(x);
		ret+="<a style='background-image:url("+imgPath+"user.gif);background-repeat:no-repeat;padding-left:18px;width:200;height:25;color:blue;' class=jellolink onclick=olItem('"+lit.Item.EntryID+"') title='"+txtClickOpen+"'><b>"+lit+"</b>";
		}
   }}catch(e){}

	 //return attachments in this action
 var lks=it.Attachments.Count;
  if (lks>0)
  {
		for (var x=1;x<=lks;x++)
		{
		var lit=it.Attachments.Item(x);
		attName=lit.DisplayName;
		if (!notEmpty(attName)){attName=txtNoSubject;}
		var lic="olpreview.gif";
		if (lit.Type==4 || lit.Type==6){lic="shcut.gif";}
		ret+="<span style='background-image:url("+imgPath+lic+");background-repeat:no-repeat;padding-left:18px;height:25;color:blue;'><a class=jellolink onclick=openAttachment('"+it.EntryID+"',"+lit.Index+") title='"+txtClickOpen+"'>"+attName+"</a>&nbsp;<a class=jellolinktop style='color:red;' onclick=delAttachment('"+it.EntryID+"',"+lit.Index+") title='"+txtDelItem+"'>[X]</a></span>&nbsp;";
		}
   }
	return ret;
}

function openAttachment(id,index)
{
//run attachment
//no way to open from Outlook still
//just open ol item for manual execution
olItem(id);
}

function delAttachment(id,index)
{
//remove an attachment
try{
  var choice=confirm(txtMsgDelItem);
  if (choice){
  var it=NSpace.GetItemFromID(id);
  it.Attachments.Remove(index);
  actioncons.innerHTML=getActionAttachmentLink(it);}
}catch(e){Ext.info.msg(txtError,txtMsgNoRmvAtt);}
}

function actionEditContacts()
{
  if (OLversion>=15)
  {alert("Cannot use this functionality in this version of Outlook");return;}
//edit contacts from the action form
var olid=Ext.getCmp("olid").getValue();
  if (!notEmpty(olid) || olid=="000")
  {//unsaved record. Save first and edit links next.
//  if(Ext.getCmp('subj').getValue()==txtNewAction){Ext.getCmp('subj').focus(true);alert(txtErrNoActionName);return;}
var g=getActiveGrid();
var noGrid=false;
if (typeof(g)=="undefined"){noGrid=true};

  var choice=confirm(txtMsgConsNeedSave);
  if (choice)
  {var naid=saveActionForm(null,noGrid,true);
  if (naid==false){return;}
  olid=naid;
  Ext.getCmp("isSaved").setValue("1");
  Ext.getCmp("olid").setValue(olid);
  actionEditContacts();
  }
  else{return;}
  }
contactList=new Array();
var recs=new Array();

  if(notEmpty(olid))
  {
  //one item
  var it=NSpace.GetItemFromID(olid);

  try{
  var lks=it.Links.Count;
  }catch(e){var lks=0;}
  //get contacts
  if (lks>0)
  {
    for (var x=1;x<=lks;x++)
    {
    var lit=it.Links.Item(x);
    contactList.push(lit.Item.EntryID);
    }
  }

  recs.push(olid);
  delegateItems(recs,true);
  }
}

function printAList(direct)
{
//print actionlist grid
var g=getActiveGrid();
if (g.getId()=="mgrid"){printMasterList(direct);return;}
var store=Ext.getCmp("grid").getStore();
var counter=0;
var ret="<head><style>.printChars{height:30px;font-size: 14px;font-family: 'Segoe UI', Verdana, 'Trebuchet MS', Arial, Sans-serif;border-bottom: 1px solid Gainsboro;}</style></head><body><table width=100% cellpadding=0 cellspacing=0>";
ret+="<tr><td class=printChars><b>Subject</b></td><td class=printChars><b>"+txtNotes+"</b></td><td class=printChars><b>"+txtDueDate+"</b></td></tr>";

  var toPrint = function(rec){
  try{var ddd=DisplayDate(rec.get("due"));}catch(e){var ddd="";}
  try{var ttt=rec.get("body");}catch(e){var ttt="";}
  try{var icls=rec.get("iclass");}catch(e){var icls="";}
  var s=ttt.search(txtTryToGetBody);
  if (s>-1 && icls<100){ttt=rec.get("notes");}else{if (icls<100){ttt+=rec.get("notes");}}
  ret+="<tr><td class=printChars><span style=font-family:Wingdings;font-size:16px;>q</span> "+rec.get("subject")+"</td><td class=printChars>"+ttt+"</td><td class=printChars>"+ddd+"</td></tr>";
  counter++;
  };

store.each(toPrint);
ret+="</table><br><br><span class=printChars>"+counter+" "+txtItemItems+"</span>";


var ff=NSpace.GetFolderFromID(jello.inboxFolder).Items;
var it=ff.Add(6);
//lastOpenTagID
try{it.Subject=lastContext+" "+txtCbActions;}catch(e){it.Subject=txtCbActions;}
it.HTMLBody=ret;
if (direct==true){it.PrintOut();Ext.info.msg(txtPrinting,txtPrintOK);}
else{it.Display();}

}



function exportToFolder(b)
{
//export all/selected items to a folder. Further Selections
if (b=="cancel"){return;}
var g=getActiveGrid();

	var ff=NSpace.PickFolder();
	if (!notEmpty(ff)){return;}
	if (ff.DefaultItemType!=3){alert(txtPromptTaskFolder);return;}

if (b=="no"){
//select all items
var rex=g.getSelectionModel().getSelections();
}
if (b=="yes"){
//select selected items
var rex=g.getStore().getRange();
}
var toFolder=ff;
var counter=0;
    for (var x=0;x<rex.length;x++){
    var r=rex[x];

    		try{
		var it=NSpace.GetItemFromID(r.get("entryID"));

		var fit=toFolder.Items.Restrict("@SQL=urn:schemas:httpmail:subject='"+it.Subject+"' AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2");
      			if (fit.Count>0)
            {
            var da=new Date(it.LastModificationTime);
            var it2=fit(1);var db=new Date(it2.LastModificationTime);
            var cmsg=txtConfirmActReplace;
            cmsg=cmsg.replace("%1",it2).replace("%2",DisplayDate(db)+"@"+DisplayTime(db)).replace("%3",it).replace("%4",DisplayDate(da)+"@"+DisplayTime(da));
            var choice=confirm(cmsg);
            if (choice==true){it2.Delete();}
            }

			var newI=it.Copy();
			var newIt=newI.Move(toFolder);
			counter++;
		}catch(e){}
    }
Ext.info.msg(txtCompleted,"<b>"+counter+"</b> "+txtItemsExported);
}


function cleanBody(bdy)
{// take a plain text body and convert to HTML for preview display
       var s = bdy;
       //;strip the img tags
       var re = /<img [^>]+>/ig;
       s = s.replace(re,"");
       // change \n to br tags
       var xe = /\n/g;
       s = s.replace(xe,"<br>\n");
       return s;
}

function appOfAction(id)
{
//create new appointment for action
try{var it=NSpace.GetItemFromID(id);}catch(e){return;}
	
	//the appointment will be created in user's calendar folder respecting settings
	if (jello.createAppointmentsOnDefCal==0 || jello.createAppointmentsOnDefCal=="0")
  {var cf=NSpace.GetFolderFromID(jello.calendarFolder);}
  else
  {var cf=calendarItems;}
  
  var t=cf.Items.Add();
  t.Subject=it.Subject;
  try{t.Attachments.Add(it);}catch(e){}
  t.itemProperties.item(catProperty).Value=it.itemProperties.item(catProperty).Value;
  t.Display();
}


function updateRecordItem(rec,removeOnly,olit)
{
//update data on ExtJS record item
var ag=getActiveGrid();
var id=rec.get("entryID");
var itype=null;
try{itype=rec.get("type");}catch(e){}
var store=ag.getStore();
var idx=store.indexOf(rec);

if (itype=="t")
{
//tag type is for remove only (actionlist specific)
store.remove(rec);
try{updateAFooterCounter(store);}catch(e){}
return;
}

//olit is tickler specific (a recurrence)
    if (olit==null)
    {

        try{var it=NSpace.GetItemFromID(id);}
        catch(e)
        {store.remove(rec);
        updateAFooterCounter(store);
        return;}

    }

    else{var it=olit;}

  var storeName=store.storeId;


	if (storeName=="tickStore")
	{	var ar=new ticklerObject(it);var newRec=new ticklerRecord(ar);}
	if (storeName=="actStore")
	{var ar=new actionObject(it);var newRec=new actionRecord(ar);}
	if (storeName=="masterStore" || storeName==null)
	{
	var tg=rec.get("tag");
	var ar=new actionObject(it);ar.tag=tg;var newRec=new actionRecord(ar);
	}
	if (storeName=="waitingStore")
	{
	var tg=rec.get("contacts");
	var ar=new actionObject(it);
	if (!notEmpty(tg))
	   {
     try{
     if (it.Links.Count>0){tg=it.Links.Item(1).Name;}else{tg="("+txtContactNo+")";}
     }catch(e){}
     }
	ar.contacts=tg;
	var newRec=new actionRecord(ar);
	}

	  if (removeOnly==false || removeOnly==null)
	  {
	  store.insert(idx,newRec);
	  }


  	  var g=ag.getSelectionModel();
  	  var wasSel=g.isSelected(rec);
  	  store.remove(rec);

	  if (wasSel==true && removeOnly!=true)
	  {var nrr=new Array();nrr.push(newRec);g.selectRecords(nrr,true);}
      try{updateAFooterCounter(store);}catch(e){}
    ag.render();
    if (idx>store.getCount()){idx=store.getCount();}
    if( jello.goNextOnDelete == false){idx--; if( idx < 0) idx=0;}
    setBodyPopDisplay(idx);
    //select next record in grid if the old one was selected
    if( wasSel)
      setTimeout(function(){
		//if (idx>store.getCount()){idx=store.getCount();}
		try{ag.getSelectionModel().selectRow(idx);ag.getView().focusRow(idx);ag.focus();}catch(e){}
		},6);

}

function getFirstAttached(att)
{
at=att.split(";");
try{var rt=NSpace.GetItemFromID(at[0]);
return rt;}catch(e){return null;}
}

function printToText()
{
//print list to not as plain text
var g=getActiveGrid();
	if (g.getId()=="grid")
	{
	//actionlist print
	var store=Ext.getCmp("grid").getStore();
	var counter=0;
	var ret="";
	ret+="["+lastContext+"] "+txtActionList+"\n\n";


	  var toTxtPrint = function(rec){
	  var cl=rec.get("iclass");
	  try{var ddd=DisplayDate(rec.get("due"));}catch(e){var ddd="";}
	  try{var ttt=rec.get("tags");}catch(e){var ttt="";}
	  try{var tbt=rec.get("body");}catch(e){var tbt="";}
	  if (cl>99){ttt="";ddd=""}

    
    if (cl<100)
	  {
    ret+="- "+rec.get("subject")+" ("+ttt+") "+ddd+" "+tbt+"\n";
    }
	  else
	  {//tag items
    ret+="- "+rec.get("subject")+" ("+tbt+")"+"\n";
    }
	    
    counter++;
	  };

	store.each(toTxtPrint);
	ret+="\n"+counter+" "+txtItemItems;

	}

	if (g.getId()=="mgrid")
	{
	//masterlist print
	var lastTag="";
	var store=Ext.getCmp("mgrid").getStore();
	var counter=0;
	var ret="";
	ret+="["+txtTaskName+" "+txtMaster+"]\n\n";

	  var toTxtMPrint = function(rec){
	  try{var ddd=DisplayDate(rec.get("due"));}catch(e){var ddd="";}
	  var tg=rec.get("tag");
	  if (lastTag!=tg){lastTag=tg;ret+="["+tg+"]\n";}
	  ret+="- "+rec.get("subject")+" "+ddd+"\n";
	  counter++;
	  };

	store.each(toTxtMPrint);
	ret+="\n"+counter+" "+txtItemItems+" "+txtOnDayWord+":"+DisplayDate(new Date());
	}

	var nte=noteItems.Items.Add();
	nte.Body=ret;
	nte.Display();
}


function saveGridState(grd)
{//save widths and showed columns of a grid
  try{
var g=Ext.getCmp(grd);
var sortOrder=null;
var st=g.getStore().getSortState();
var sfld=st.field;var sord=st.direction;
var columns=new Array();
g.getColumnModel().getColumnsBy(function(c){
	var sortOrder=null;
	if(c.dataIndex==sfld){sortOrder=sord;}
	columns.push({
	  	hidden: c.hidden,
	  	width: c.width,
	  	dataIndex:c.dataIndex,
	  	sorted:sortOrder
	});
});
var coldata=Ext.encode(columns);

var idx=columnStore.find("grid",new RegExp("^"+grd+"$"));
  if (idx>-1)
  {
  var r=columnStore.getAt(idx);
  r.beginEdit();
  r.set("data",coldata);
  r.endEdit();
  }
  else
  {
  var tr=new columnRecord({grid:grd,data:coldata});
  columnStore.add(tr);
  }
syncStore(columnStore,"jello.gridColumns");
jese.saveCurrent();
  }catch(e){}
}

function restoreGridState(grd)
{//restore saved column widths and hidden status of a grid
var g=Ext.getCmp(grd);
var store=g.getStore();
var idx=columnStore.find("grid",new RegExp("^"+grd+"$"));
  if (idx==-1){return;}
  var r=columnStore.getAt(idx);
  var newConfig=Ext.decode(r.get("data"));
  var oldCol;
  var doSortFld=null;var doSortOrd=null;
  var cm = g.getColumnModel();
      for (var i = 0; i < newConfig.length; i++) {

      var ncdi=newConfig[i].dataIndex;
      if (typeof(ncdi)!="undefined")
      {
      oldCol=cm.findColumnIndex(newConfig[i].dataIndex);
      if (oldCol != i) {
      cm.moveColumn(oldCol, i); // Order
      }
      }
          try{cm.setHidden(i, newConfig[i].hidden);}catch(e){} //Hidden
          cm.setColumnWidth(i, newConfig[i].width); // Widths
          try{
          if (newConfig[i].sorted!=null){doSortFld=newConfig[i].dataIndex;doSortOrd=newConfig[i].sorted;}
          }catch(e){}
      }

   try{store.sort(doSortFld,doSortOrd);}catch(e){}
}

function getActionFormToolbar(fi,toForm)
{
var sufx="frm-";
if (toForm){sufx;}
//toolbar for the action form
        tbar=[{
        icon: 'img\\new.gif',
        cls:'x-btn-icon',
        id:sufx+'anew',
        tooltip: txtAddNewItmInfo,
		handler : actionFormSelected
        },{
        icon: 'img\\page_delete.gif',
        cls:'x-btn-icon',
        id:sufx+'adel',
        tooltip: txtDelItmInfo,
        handler : actionFormSelected
        },'-',{
        icon: 'img\\IsNext.gif',
        cls:'x-btn-icon',
        id:sufx+'anext',
        tooltip: txtSetNAInfo,
        handler : actionFormSelected
        },{
        icon: 'img\\check.gif',
        cls:'x-btn-icon',
        id:sufx+'adone',
        tooltip: txtCompleteItmInfo,
        handler : actionFormSelected
        },
        '-',{
                tooltip: txtsetPriority,
                cls:'x-btn-icon',
                id:sufx+'apri',
                icon: 'img\\priority.gif',
                handler: actionFormSelected
        },{
                tooltip: txtStartStop,
                cls:'x-btn-icon',
                id:sufx+'astart',
                icon: 'img\\widget-add.gif',
                handler:actionFormSelected
        },{
                tooltip: txtSetPrivacy,
                cls:'x-btn-icon',
                id:sufx+'alock',
                icon: 'img\\page_lock.gif',
                handler:actionFormSelected
        },{
                tooltip: txtOpenInOutlook,
                cls:'x-btn-icon',
                id:sufx+'aoutlook',
                icon: 'img\\info.gif',
                handler:actionFormSelected
        },'-',{
        icon: 'img\\appoint.gif',
        cls:'x-btn-icon',
        tooltip: txtItmAppInfo,
        id:sufx+'aappoint',
        handler : actionFormSelected
        },{
        icon: 'img\\shcut.gif',
        cls:'x-btn-icon',
        id:sufx+'ashort',
        tooltip: txtShcutInfo,
        handler : actionFormSelected
        },'-',
        {
        icon: 'img\\page_prev.gif',
        cls:'x-btn-icon',
        id:sufx+'agnotes',
        tooltip: txtActFANotesfromOL,
        handler : actionFormSelected
        },{
        icon: 'img\\page_next.gif',
        cls:'x-btn-icon',
        id:sufx+'asnotes',
        tooltip: txtActFANotestoOL,
        handler : actionFormSelected
        },'->',fi

        ];

return tbar;
}

//new
function actionFormSelected(btn)
{
//handle toolbar buttons of action form
var id=Ext.getCmp("olid").getValue();
var it=NSpace.getItemFromID(id);
var aFailed=false;
var delOnly=false;

if (btn.id.substr(0,4)=="frm-"){btn.id=btn.id.substr(4,10);}

	 //save and new
	 if (btn.id=="anew")
   {saveActionForm(null,true);
   setTimeout(function(){editAction(null,true,false);},3);
   }
	 //delete
     if (btn.id=="adel")
		{
		var choice=confirm(txtMsgDelItem);
		if (choice){OLDeleteItem(id);Ext.getCmp("theactionform").destroy();delOnly=true;}
		}
     //toggle NA
     if (btn.id=="anext")
		{aFailed=doAction(id);updateActFormDisplay(it);
		}
     //toggle Done
     if (btn.id=="adone")
		{
			if (it.Complete==false)
			{
			var choice=confirm(txtADone+"?");
			if (!choice){return;}
			}
		aFailed=actionDone(id);
		var comp=it.Complete;var xt="";
		if (comp){Ext.getCmp("theactionform").destroy();}
		else{alert(txtAction+" "+txtUncompleted);}
		}
     //toggle Priority
     if (btn.id=="apri")
		{aFailed=setActionPriority(id);
		var pp=getJPriorityProperty(it);
		Ext.getCmp("pri").setValue(pp);
		}
     //toggle Start/stop
     if (btn.id=="astart")
		{
		aFailed=startStopAction(id);
		if (it.Status==0){alert(txtAction+":"+txtNotStarted);}
		else{alert(txtAction+":"+txtInProgress);}
		}
     //toggle Privacy
     if (btn.id=="alock")
		{
		aFailed=lockAction(id);
		if (it.Sensitivity==2){alert(txtAction+":"+txtMarkContextPrivate);}
		else{alert(txtAction+":Non Private");}
		}
     //open in outlook
     if (btn.id=="aoutlook"){olItem(id);}
     //create appointment
     if (btn.id=="aappoint"){appOfAction(id);}
     //create shortcut
     if (btn.id=="ashort"){alShortcut(id,"a");}
	//get outlook notes
	if (btn.id=="agnotes")
		{
		 var nf=Ext.getCmp("notes").getRawValue();
			if (notEmpty(nf))
			{
			var choice=confirm(txtPrmptReplAcNotes);
			if (!choice){return;}
			}
			try{var of=it.Body;}catch(e){return;}
			Ext.getCmp("notes").setValue(of);
		}
	//set outlook notes
	if (btn.id=="asnotes")
		{
		 var nf=Ext.getCmp("notes").getRawValue();
		try{var of=it.Body;}catch(e){return;}
			if (notEmpty(of))
			{
			var choice=confirm(txtPrmptReplOLNotes+"\n\n"+of.substr(0,100)+"...");
			if (!choice){return;}
			}
			it.Body=Ext.getCmp("notes").getRawValue();
			it.Save();
			alert(txtActionSuccess);
		 }



//update grid record if there is one
var g=getActiveGrid();
	if (typeof(g)!="undefined")
	{
		var r=g.getSelectionModel().getSelected();
		if (r!=null){
			var eid=r.get("entryID");
			if (btn.id=="adel" || eid==it.EntryID)
			{updateRecordItem(r,delOnly,null);}}

	}

}

function updateActFormDisplay(it)
{
//update tag list
var ftagList=it.itemProperties.item(catProperty).Value+";";
ftagList=getUniqueTags(ftagList);
var  tgs=tagList("formtagdisplay",null,ftagList,it.EntryID);
Ext.getCmp("formtagdisplay").getEl().update(tgs[0]);
Ext.getCmp("tagfield").setValue(ftagList);
Ext.info.msg(txtActionEdit,txtActionSuccess);
}

function checkForExistingGridToUpdate(grd,it)
{
	if (typeof(grd)=="undefined"){return;}
	if (!notEmpty(lastContext)){return;}
	var ftagList=it.itemProperties.item(catProperty).Value+";";
	ftagList=getUniqueTags(ftagList);
	var o=ftagList.split(";");
	var relevantGrid=false;
		if (o.length>0)
		{
			for (var x=0;x<o.length;x++)
			{
				if (lastContext==o[x]){relevantGrid=true;break;}
			}
		}

	if (grd.getId().substr(0,3)=="olv"){relevantGrid=false;}
	if (relevantGrid)
	{
	//relevant grid found open. update/Create item
	var ar=new actionObject(it);var newRec=new actionRecord(ar);
	var store=grd.getStore();
	store.add(newRec);
	var nrr=new Array();nrr.push(newRec);grd.getSelectionModel().selectRecords(nrr,true);
      try{updateAFooterCounter(store);}catch(e){}
	}
}

function printWithTicklersForm()
{
getTagsArray();
var dts=jello.dateSeperator;

var dccformat="n"+dts+"j"+dts+"Y";
var dccalt="m/d|n/j|m-d|n-j|m/d/Y|m/d/Y|m/d/y|n/j/y|j/n/y";
if(jello.dateFormat == 0 || jello.dateFormat =="0"){dccformat="j"+dts+"n"+dts+"Y";dccalt="d/m|j/n|d-m|j-n|d/m/Y|d/m/y";}
if(jello.dateFormat == 2 || jello.dateFormat =="2"){dccformat="Y"+dts+"m"+dts+"d";dccalt="Y/m/d|y/m/d|Y-m-d";}

var tt=new Date();
fdue=tt;
var wfos=jello.weeklyListDefs.split("|");

 var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        height:420,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'wpform',
        //iconCls:'tagformicon',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',

        items: [
				new Ext.form.DateField({
                fieldLabel: txtWeekSelector,
                id: 'week',
                format:dccformat,
				invalidText:txtInvDate,
				width:100,
				altFormats:dccalt,
                value:fdue,
                listeners:
                {
                change:function(f,nv){var rt=wpWeekInfo(nv);Ext.getCmp("weekinfo").getEl().update(rt);}
                }
            }),
				new Ext.form.Label({
				html:wpWeekInfo(fdue),
				height:40,
				id:'weekinfo'
				}),

				{
                fieldLabel: txtTag,
                id: 'ftag01',
                value:lastContext,
                disabled:true,
                allowBlank:false
            }
				,
			new Ext.form.NumberField({
			fieldLabel:txtCalBoxHeight,
			value:wfos[0],
			width:200,
			invalidText:txtInvalidTxt,
			hideLabel:false,
			id:'calboxh'
	}),	new Ext.form.NumberField({
			fieldLabel:"Limit Body Text Display Characters (0 means all)",
			value:0,
			width:200,
			invalidText:txtInvalidTxt,
			hideLabel:false,
			id:'bodylimit'
	}),
	new Ext.form.NumberField({
			fieldLabel:txtBlankLinesNum,
			value:wfos[1],
			width:200,
			invalidText:txtInvalidTxt,
			hideLabel:false,
			id:'blanks'
	}),	{
                fieldLabel: txtBlanksTitle,
                id: 'blkttl',
                value:wfos[2],
                allowBlank:true
            },
                new Ext.form.ComboBox({
                fieldLabel: txtAdditionalTag+' #1',
                id: 'addtag',
                store:globalTags,
                triggerAction:'all',
                hideTrigger:false,
                typeAhead:true,
                value:wfos[3],
                mode:'local'
	            }),
	             new Ext.form.ComboBox({
                fieldLabel: txtAdditionalTag+' #2',
                id: 'addtag2',
                store:globalTags,
                triggerAction:'all',
                hideTrigger:false,
                typeAhead:true,
                value:wfos[4],
                mode:'local'
	            })
        ]

    });



   var win = new Ext.Window({
        title: txtPtSndWTicklers,
        width: 480,
        height:480,
        id:'thewpform',
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
            text:txtPrint,
            id:'sbut',
            tooltip:txtPrint+' [F2]',
            listeners:{
            click: function(b,e){
            printWithTicklers();
            }
            }
        },{
            text: 'Clear Tags',
            tooltip:'Clear Additional Tags',
                        listeners:{
            click: function(b,e){
            Ext.getCmp("addtag").setValue("");
            Ext.getCmp("addtag2").setValue("");

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
  setTimeout(function(){Ext.getCmp("week").focus(true,100);},1);

//shortcut keys
var map = new Ext.KeyMap('thewpform', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});

var map2 = new Ext.KeyMap('thewpform', {
    key: Ext.EventObject.F2,
    fn: function(){
    try{
    printWithTicklers();}catch(e){}},
    scope: this});
}

function wpWeekInfo(fdue)
{
if (fdue==null){fdue=Ext.getCmp("week").getValue();}
	var now =  fdue;
	var nowDayOfWeek = now.getDay()-1;var nowDay = now.getDate();
	var nowMonth = now.getMonth();var nowYear = now.getYear();
	nowYear += (nowYear < 2000) ? 1900 : 0;
	dsta = new Date(nowYear, nowMonth, nowDay - nowDayOfWeek);
	dend = new Date(nowYear, nowMonth, nowDay + (6 - nowDayOfWeek));

return "<span style=width:100px;>&nbsp;</span><span class=tagprj>"+txtWeek+" "+fdue.format('W')+"</span> "+DisplayAppDate(dsta)+" - "+DisplayAppDate(dend);
}

function printWithTicklers()
{
//Create a printout with this list's items + Weekly view and Waiting list

var g = getActiveGrid();
var store = g.getStore();
var counter=0;
var firstContext=lastContext;
var ret="<head><style>ul{list-style-type:circle;}body{font-size: 14px;font-family: 'Segoe UI', Verdana, 'Trebuchet MS', Arial, Sans-serif;}.printChars{height:20px;border-bottom: 1px solid Gainsboro;}.calbox{border:1px solid Gainsboro;font-size:10px;}.calbox2{border:1px solid Gainsboro;background:gainsboro;font-size:10px;}.listbox{font-size:10px;height:20px;border-left:1px solid Gainsboro;border-right:1px solid Gainsboro;border-bottom:1px solid Gainsboro;}.listbox2{height:20px;background:gainsboro;border-left:1px solid Gainsboro;border-right:1px solid Gainsboro;border-bottom:1px solid Gainsboro;}</style></head><body><table width=95% cellpadding=0 cellspacing=0>";
var nd=Ext.getCmp("week").getValue();
var bodyLimit= Ext.getCmp("bodylimit").getValue();

	var now =  nd;
	var nowDayOfWeek = now.getDay()-1;var nowDay = now.getDate();
	var nowMonth = now.getMonth();var nowYear = now.getYear();
	nowYear += (nowYear < 2000) ? 1900 : 0;
	var dsta = new Date(nowYear, nowMonth, nowDay - nowDayOfWeek);
	var dend = new Date(nowYear, nowMonth, nowDay + (6 - nowDayOfWeek));
    var xdsta = new Date(nowYear, nowMonth, nowDay - nowDayOfWeek);
    var xdend = new Date(nowYear, nowMonth, nowDay + (6 - nowDayOfWeek));
//get ticklers
var tStore = new Ext.data.SimpleStore({
fields: [
    {name: 'subject'},
    {name: 'type'},
    {name: 'entryID'},
	{name: 'isrecurring',type:'boolean'},
	{name: 'attachment'},
	{name: 'importance'},
	{name: 'created',type:'date'},
	{name: 'iclass'},
	{name: 'contacts'},
	{name: 'daypos'},
	{name: 'sensitivity'},
	{name: 'start'},
	{name: 'status'},
	{name: 'location'},
	{name: 'due',type:'date'},
	{name: 'end',type:'date'},
	{name: 'reminder',type:'date'},
	{name: 'unread',type:'boolean'},
	{name: 'allday',type:'boolean'},
	{name: 'duration'},
	{name: 'done'},
	{name: 'tags'}
        ]
    });

getTicklerFile(dsta,dend,1,tStore,true,true);

var cp=new Array(6);

//loop returned ticklers
  var toCalendar = function(rec){
	var ds=rec.get("due").getDay();
	if (!notEmpty(cp[ds])){cp[ds]="<ul>";}
		var tmm="**" + txtAllDay + "** - ";
        cl = rec.get("iclass");
        if( cl == 43 || cl == 48)
            tmm = "";
        // if allday not start time should be displayed
		if( rec.get("allday") == false && cl != 43 && cl != 48){
            try{
            var sd=new Date(rec.get("start"));
            tmm=sd.format("H:i")+"&nbsp;";
            if (jello.timeFormat==1 || jello.timeFormat=="1"){tmm=sd.format("g:i")+"&nbsp;";}
            }catch(e){}
        }
	cp[ds]+="<span class=printchars><li>"+tmm+rec.get("subject")+"</span>";
  };
tStore.each(toCalendar);
var hgt=Ext.getCmp("calboxh").getValue();
if (!notEmpty(hgt)){hgt=80;}
//render calendar
var thisday;
var daylist=txtDayList.split(",");
ret+="<table width=100% cellpadding=0 cellspacing=0>";
ret+="<tr><td colspan=2 ALIGN=CENTER ><span class=printchars><b>"+DisplayAppDate(dsta)+" - "+DisplayAppDate(dend)+"</b></td></tr>";
thisday=dsta.getDate();
ret+="<tr><td class=calbox2 WIDTH=50%><span class=printchars><b>"+daylist[1]+" "+thisday+"</b></td>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<td class=calbox2 WIDTH=50%><span class=printchars><b>"+daylist[2]+" "+thisday+"</b></td></tr>";
var cpt="";if (notEmpty(cp[1])){cpt=cp[1];}
ret+="<tr><td class=calbox VALIGN=TOP height="+hgt+">"+cpt+"</td>";
cpt="";if (notEmpty(cp[2])){cpt=cp[2];}
ret+="<td class=calbox>"+cpt+"</td></tr>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<tr><td class=calbox2><span class=printchars><b>"+daylist[3]+" "+thisday+"</b></td>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<td class=calbox2><span class=printchars><b>"+daylist[4]+" "+thisday+"</b></td></tr>";
cpt="";if (notEmpty(cp[3])){cpt=cp[3];}
ret+="<tr><td class=calbox VALIGN=TOP height="+hgt+">"+cpt+"</td>";
cpt="";if (notEmpty(cp[4])){cpt=cp[4];}
ret+="<td class=calbox VALIGN=TOP>"+cpt+"</td></tr>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<tr><td class=calbox2><span class=printchars><b>"+daylist[5]+" "+thisday+"</b></td>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<td class=calbox2><span class=printchars><b>"+daylist[6]+" "+thisday+"</b></td></tr>";
cpt="";if (notEmpty(cp[5])){cpt=cp[5];}
ret+="<tr><td class=calbox VALIGN=TOP>"+cpt+"</td>";
cpt="";if (notEmpty(cp[6])){cpt=cp[6];}
ret+="<td class=calbox VALIGN=TOP><table width=100%><tr>";
ret+="<td class=calbox VALIGN=TOP height="+(hgt/2)+" style=border:none;>"+cpt+"</td></tr>";
dsta.setDate(dsta.getDate()+1);
thisday=dsta.getDate();
ret+="<tr><td class=calbox2 width=100%><span class=printchars><b>"+daylist[0]+" "+thisday+"</b></td></tr>";
cpt="";if (notEmpty(cp[0])){cpt=cp[0];}
ret+="<tr><td VALIGN=TOP class=calbox style=border:none;>"+cpt+"</td></tr></table>";
ret+="</table>";
ret+="<table width=95% cellpadding=0 cellspacing=0>";
//list
ret+="<tr><td colspan=2 class=listbox2><b>"+lastContext+"</b></td></tr>";
var col=0;

  var toPrint = function(rec){
  try{var ddd="("+DisplayDate(rec.get("due"))+")";}catch(e){var ddd="";}
  if (ddd=="()"){ddd="";}
	try{var ttt=rec.get("body");}catch(e){var ttt="";}
	var s=ttt.search(txtTryToGetBody);
	if (s>-1){ttt=rec.get("notes");}else{ttt+=rec.get("notes");}
  if( bodyLimit != 0) {ttt = ttt.replace(/\n/g,' ');ttt = ttt.substring(0,bodyLimit);}
  if (col==0){ret+="<tr>";}
  var ccc=" ("+rec.get("tags").replace(lastContext,"")+")";if (ccc==" ()"){ccc="";}
  ret+="<td width=50% class=listbox><span style=font-family:Wingdings;font-size:16px;>q</span> "+rec.get("subject")+ccc+" "+ttt+" "+ddd+"</td>";
  if (col==1){ret+="</tr>";col=-1;}
  counter++;
  col++;
  };



store.each(toPrint);

if (col==1){ret+="<td class=listbox>&nbsp;</td></tr>";}

//add blank lines
var bl=Ext.getCmp("blanks").getValue();
if (bl>0)
{
var bt=Ext.getCmp("blkttl").getValue();
if (notEmpty(bt))
{ret+="<tr><td colspan=2 class=listbox2><b>"+bt+"</b></td></tr>";}

for (var x=0;x<bl;x++)
{ret+="<tr><td width=50% class=listbox>&nbsp;</td><td class=listbox>&nbsp;</td></tr>";}
}

//add additional tags

  var toExtraPrint = function(rec){
  try{var ddd="("+DisplayDate(rec.get("due"))+")";}catch(e){var ddd="";}
  if (ddd=="()"){ddd="";}
  try{var ttt=rec.get("body");}catch(e){var ttt="";}
  var s=ttt.search(txtTryToGetBody);
  if (s>-1){ttt=rec.get("notes");}else{ttt+=rec.get("notes");}

  if( bodyLimit != 0) {ttt = ttt.replace(/\n/g,' ');ttt = ttt.substring(0,bodyLimit);}
  ret+="<li>"+rec.get("subject")+" "+ttt+" "+ddd;
  };

	var t2=Ext.getCmp("addtag").getValue();
	var t3=Ext.getCmp("addtag2").getValue();
	ret+="<tr><td class=listbox2><b>"+t2+"</b></td>";
	ret+="<td class=listbox2><b>"+t3+"</b></td></tr><tr><td class=listbox valign=top><ul>";

	if (notEmpty(t2))
	{
	var t2store=showContext(t2,2);
	t2store.each(toExtraPrint);
	ret+="</td>";
	}
	else {ret+="&nbsp;</td>";}

	ret+="<td class=listbox valign=top><ul>";

	if (notEmpty(t3))
	{
	var t2store=showContext(t3,2);
	t2store.each(toExtraPrint);
	ret+="</td>";
	}else {ret+="&nbsp;</td>";}

	ret+="</tr>";

var ff=NSpace.GetFolderFromID(jello.inboxFolder).Items;
var it=ff.Add(6);
//lastOpenTagID
var stdatedisp = DisplayAppDate(xdsta);
var enddatedisp =DisplayAppDate(xdend);
stdatedisp = stdatedisp.replace(/\<span[^\>]*\>/,"");
stdatedisp = stdatedisp.replace(/\<\/span\>/,"");
enddatedisp = enddatedisp.replace(/\<span[^\>]*\>/,"");
enddatedisp = enddatedisp.replace(/\<\/span\>/,"");
try{it.Subject=firstContext+" "+txtCbActions+":"+txtWeek + "  " +stdatedisp+" - "+enddatedisp;}catch(e){it.Subject=txtCbActions;}
it.HTMLBody=ret;
it.Display();
//save settings for later
if (!notEmpty(bt)){bt=txtNotes;}
jello.weeklyListDefs=hgt+"|"+bl+"|"+bt+"|"+t2+"|"+t3;
jese.saveCurrent();
//destroy form
Ext.getCmp("thewpform").destroy();
if( g.getId() == "mgrid")
setTimeout(
function(){pMaster();
},6);
else
setTimeout(
function(){showContext(firstContext,1);
},6);
}



function popTaskBody(id,rec)
{
if (jello.expBodyView=="1" || jello.expBodyView==1){popTaskBodyExperimental(id, rec);return;}

	var it=NSpace.GetItemFromID(id);
	var init_t;
	var init_typ;
	var isIGrid=null;
  try{
  isIGrid=getActiveGrid().getId();
  }catch(e){}
	if (jello.mailPreviewHTML==true){init_t=it.HTMLBody;init_typ="html";}
	else{init_t=cleanBody(it.Body);init_typ="text";}
	var bpanel = new Ext.form.TextArea({
		id: 'bpanel',
		value:it.Body,
		height: 340,
		width: 450,
		readOnly: false,
		disabled: (isIGrid == "igrid"? true:false),
		hidden: (isIGrid  == "igrid"? true:false),
		idval: init_t,
		viewval: "text"
	});
	var init_t;
	var init_typ;
	if (jello.mailPreviewHTML==true){init_t=it.HTMLBody;init_typ="html";}
	else{init_t=cleanBody(it.Body);init_typ="text";}
var win = new Ext.Window({
        title: "Body Text",
        width: 480,
        height:370,
        id:'bodytxt',
        minWidth: 300,
        minHeight: 280,
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        //tbar: tbr,
        items: [ 
		{
             html:init_t,
             id:'bodyval',
             autoWidth:true,
             region:'center',
             border:false,
             autoScroll:true,
             disabled: (isIGrid == "igrid"? false:true),
			 hidden: (isIGrid == "igrid"? false:true),
             bodyStyle: {
             background: '#ffffff',
             padding: '0px'
             },
             layout:'anchor'
         },
		 bpanel,
		 new Ext.form.Hidden({
				id:'olid',
				value:id
				})],
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
	        text: '<b>'+txtSave+'</b>',
	        disabled: (isIGrid == "igrid"? true:false),
			hidden: (isIGrid == "igrid"? true:false),
            listeners:{
            click: function(b,e){
		        	var tid=Ext.getCmp("olid").getValue();
					var it=NSpace.GetItemFromID(tid);
					var x = Ext.getCmp("bpanel");
					if( x.isDirty()){
						it.Body = x.getRawValue();
						it.Save();
						var gr = getActiveGrid();
						var store = gr.getStore();
						var recno = store.find("entryID",tid);
						if( recno != -1){
							rec = store.getAt(recno);
							rec.set("body", getItemBody(it));
						}
						gr.getView().refresh();
						win.destroy();
					}
				}
            }
        },{
			text: "Toggle View",
			hidden : ((it.Class == 43  && (isIGrid == "igrid"))?false:true),
			listeners:{
				click: function(b,e){
					var g = Ext.getCmp("bodytxt");
					var tid= Ext.getCmp("olid").getValue();
					var it=NSpace.GetItemFromID(tid);
					if( getActiveGrid().getId() != "igrid")
						return;
	 				var x = Ext.getCmp("bodyval");
 
					if(g.viewval == "text")	 {
						x.body.dom.innerHTML=it.HTMLBody;g.viewval = "html";

					}else{
						x.body.dom.innerHTML=cleanBody(it.Body);g.viewval = "text";
					}
					
				}
			}
		},
		{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
	        listeners:{
		        click: function(b,e){
            win.destroy();}
			}
        }
	
		
		
		],
 		viewval: init_typ
    });
 
    showOLViewCtl(false);
     win.show();
    win.setActive(true);
    var map = new Ext.KeyMap('bodytxt', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});
	
}
var lastView = null;
function setActionView()
{
 	var g = getActiveGrid();
 	if( g.getId() == 'grid')
		var iF=NSpace.GetFolderFromID(jello.actionFolder);
	else{
		var fldx = lIbFolder;
		if( fldx == "" || fldx == null)
			ctx = jello.inboxFolder;
		var iF=NSpace.GetFolderFromID(ctx);
	}
 	var views = iF.Views;
 	if(views.Count == 0)
 		return;
 	var varray = new Array();
 	for( var i=1; i <= views.Count; i++)
 	{
 		try{var vw = views.Item(i);
 		varray.push(new Array(vw.name, vw.Filter));
 		}catch(e){continue;}
 	}
 	vnameStore = new Ext.data.SimpleStore({
				fields: ['name','filter'],
				data : varray
			});
	var combo = new Ext.form.ComboBox({
			id: "selActionView",
			store: vnameStore,
			anyMatch: true,
			matchWordStart: false,
			queryIgnoreCase: true,
			displayField:'name',
			typeAhead: true,
			mode: 'local',
			//emptyText:txtGotoInfo+'...',
			selectOnFocus:true,
			resizable: true,
			minListWidth: 110,	
			listeners : {
				select : function(combo,record,idx){
					
					var iF=NSpace.GetFolderFromID(jello.actionFolder);
					var g = getActiveGrid();
					var ctx = g.ctx;
					if( g.getId() == "grid"){
					var flt= record.get("filter");
					flt = (flt==""?"":"("+flt+") AND ") +
						"(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "')";
          			showContext(ctx,1,flt,4);
          			}else{
          				ctx = lIbFolder;
						if( ctx == "" || ctx == null)
          					ctx = jello.inboxFolder;
          				var iF=NSpace.GetFolderFromID(ctx);
          				var flt = record.get("name");
          				var vws = iF.Views;
          				var v = vws.Item(flt);
          				v.Apply();
          				pInbox(lIbFolder);
          			}
					setTimeout(function(){
        			var w = Ext.getCmp("selectview");	
							w.close();},6);	
				}
			}
		});
	var win = new Ext.Window({
        title: "Select View",
        width: 300,
        height: 100,
        id:'selectview',
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: combo,
        modal:true,
        listeners: {destroy: function(t){showOLViewCtl(true);}},
        buttons: [
			{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();}
            }}
		
        ]
    });
	
    showOLViewCtl(false);
	win.show();

    win.setActive(true);
  setTimeout(function(){Ext.getCmp("selActionView").focus(true,100);},1);
  	
	//shortcut keys
var map = new Ext.KeyMap('selActionView', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});
 }
 
function popTaskBodyExperimental(id, rec)
{
    var init_t;
	var init_typ;
    var strid="";
    var recno=null;
    var aGrid = getActiveGrid();
    var aGridID = aGrid.getId();
    var isIgrid = (aGridID == "igrid"?true:false);
    var xWidth=(Ext.getCmp("viewport").getWidth()*75)/100;
    var xHeight = (Ext.getCmp("viewport").getHeight()*70)/100;
	var pWidth = (xWidth< 480?480:xWidth-20);
	var pHeight = (xHeight< 460?280:(xHeight-200));

    if( id == null && isIgrid)
    {
        rec = aGrid.getSelectionModel().getSelected();
        id = rec.get("entryID");

    }
	// if no record, get the current rec if a grid we know about
	if( rec == null){
		if( aGridID == "grid" || aGridID == "tgrid" || aGridID == "mgrid")
			rec = aGrid.getSelectionModel().getSelected();
	}
	
	try{var it=NSpace.GetItemFromID(id);} 
	catch(e)
	{var x = rec.store;
	x.remove(rec);aGrid.render();return;
	}

    var oWin = Ext.getCmp("bodytxt");
    if( oWin != null && typeof(oWin) != "undefined"){
        if( isIgrid)
            oWin.destroy();
        else{
            oWin.show();
            var idx = recno = rec.store.indexOf(rec);
            setBodyPopDisplay(idx);
            return;
        }
    }

    if( rec != null && typeof(rec) != "undefined")
    {
        strid = rec.store.id;
        recno = rec.store.indexOf(rec);
    }
	var foldercombo;
	if( isIgrid){
		buildFolderList(false,false);
		foldercombo = new Ext.form.ComboBox({
			id: "popFoldCombo",
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
			folId: "",
			forceSelection: true,
			allowBlank: true,
			disabled: (isIgrid? false:true),
			hidden: (isIgrid? false:true),
			listeners : {
				select : function(combo,record,idx){
					var id = record.get("id");
					Ext.getCmp("popFoldCombo").folId = id;
				}
			}

		});
	}
	if (jello.mailPreviewHTML==true){init_t=it.HTMLBody;init_typ="html";}
	else{init_t=cleanBody(it.Body);init_typ="text";}
	var bpanel = new Ext.form.TextArea({
		id: 'bpanel',
        fontsize: jello.mailPreviewFontSize,
		value:it.Body,
        width: pWidth,
        height:pHeight,
        border: true,
		readOnly: false,
        autoScroll: true,
        margins: {bottom:10},
		disabled: (isIgrid? true:false),
		hidden: (isIgrid? true:false),
		idval: init_t,
		viewval: "text"
	});

	var tgs="";var ftagList="";
	if( isIgrid){
    // now add tagging area

    tgs=tagList('pformtagdisplay',null,it.itemProperties.item(catProperty).Value);
    ftagList=it.itemProperties.item(catProperty).Value;
    if(tgs==""){tgs=" ";}
    var cbox2 = new Ext.form.ComboBox({
        fieldLabel: txtTags,
        id: 'ptags',
        store:globalTags,
        height: 40,
        width: 320,
        autoWidth: true,
        hideTrigger:false,
        typeAhead:true,
        mode:'local',
        disabled: (isIgrid? false:true),
		hidden: (isIgrid? false:true),
        listeners:{
        blur:function(){quickAddTag("ptags", "ptagfield",'pformtagdisplay',Ext.getCmp("popolid").getValue());status=txtReady;},
        specialkey: function(f, e){
        if(e.getKey()==e.ENTER){
        userTriggered = true;
        e.stopEvent();
        f.el.blur();
        quickAddTag("ptags", "ptagfield",'pformtagdisplay',Ext.getCmp("popolid").getValue());
        status=txtReady;
        setPopCategory();}},

        select: function(cb,rec,idx){
        quickAddTag("ptags", "ptagfield",'pformtagdisplay',Ext.getCmp("popolid").getValue());
        status=txtReady;
        setPopCategory();
        }
        }

    });

    var catdisp = new Ext.form.Label({
    html:tgs[0],
    height:50,
    width: 450,
    autoWidth: true,
    disabled: (isIgrid? false:true),
	hidden: (isIgrid? false:true),
    //cls:'formtaglist',
    id:'pformtagdisplay'
    });


    var cats = new Ext.form.Hidden({
    id:'ptagfield',
    value:ftagList
    });

	var hdisp = {
             html:"<font style='font-size: " + jello.mailPreviewFontSize +"px'>"+ init_t + "</font>",
             id:'bodyval',
			width: pWidth,
			height:pHeight,
             region:'center',
             border:true,
             margins: {bottom:10},
             autoScroll:true,
             disabled: (isIgrid? false:true),
			 hidden: (isIgrid? false:true),
             bodyStyle: {
             background: '#ffffff',
             padding: '0px'
             },
             layout:'anchor'
         };
	}
var tbar = popToolbar(it,isIgrid);
var iArray = new Array();
if( isIgrid)
	iArray.push(hdisp);
iArray.push(bpanel);
if(isIgrid){
iArray.push(foldercombo);
iArray.push(cbox2);
iArray.push(catdisp);
iArray.push(cats);
}
var win = new Ext.Window({
        title: it.Parent.FolderPath + "(" + (recno+1).toString() +"/" + it.Parent.Items.Count +")",
        width: (xWidth< 480?480:xWidth),
        height:(xHeight< 460?460:xHeight),
        id:'bodytxt',
        tbar: tbar,
        minWidth: 480,
        minHeight: 460,
        margins: {bottom:5},
        resizable:true,
        draggable:true,
        layout: (isIgrid?'form':'auto'),
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        modal:(isIgrid?false:true),
        minimizable: (isIgrid?true:false),
 		viewval: init_typ,
        // this are defined for use by variuous funcitons and aren't
        // part of a normal window
        items: [
		iArray,

		 new Ext.form.Hidden({
				id:'popolid',
				value:id
				}),
        new Ext.form.Hidden({
				id:'grecno',
				value:recno
				})],

        listeners: {
            destroy:function(t){
            showOLViewCtl(true);
            //setTimeout(function(){refreshView();},50);
            },
            minimize: function(t){win.hide();}
        },
        buttons: []
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);
    if( isIgrid)
		Ext.getCmp("bodyval").body.setStyle("font-size",jello.mailPreviewFontSize);
    setTimeout(function(){setBodyPopDisplay(recno);},6);
    var map = new Ext.KeyMap('bodytxt', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});

}

function popToolbar(it, isIgrid)
{
    var tbar = [
    {
            icon:'img\\home16.png',
            cls:'x-btn-icon',
            tootip: txtHome,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
            listeners:{
				click: function(b,e){
                    setBodyPopDisplay(0);
                }
            }
        },{
            icon:'img\\page_prev.gif',
            cls:'x-btn-icon',
            tootip: txtPrevious,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
            listeners:{
				click: function(b,e){
                    goBodyPopMessage(-1);
                }
            }
        },{
            icon:'img\\page_next.gif',
            cls:'x-btn-icon',
            tootip: txtNext,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
            listeners:{
				click: function(b,e){
                    goBodyPopMessage(1);
                }
            }
        },'-',{
	        text: '<b>'+txtSave+'</b>',
	        disabled: (isIgrid? true:false),
			hidden: (isIgrid? true:false),
            listeners:{
            click: function(b,e){
		        	var tid=Ext.getCmp("popolid").getValue();
					var it=NSpace.GetItemFromID(tid);
					var x = Ext.getCmp("bpanel");
					if( x.isDirty()){
						it.Body = x.getRawValue();
						it.Save();
						var gr = getActiveGrid();
						var store = gr.getStore();
						var recno = store.find("entryID",tid);
						if( recno != -1){
							rec = store.getAt(recno);
							rec.set("body", getItemBody(it));
						}
						gr.getView().refresh();
						var win= Ext.getCmp("bodytxt");
						win.destroy();
					}
				}
            }
        },{
        icon:'img\\move.gif',
        cls:'x-btn-icon',
		hidden : ((it.Class == 43  && (isIgrid))?false:true),
        tooltip: txtAToFolder,
		listeners:{click: function(b,e){
            var fdest = Ext.getCmp("popFoldCombo").folId;
            if( fdest == "" || fdest == null){
                (function(){ Ext.getCmp('popFoldCombo').focus();}).defer(100);
                return;
            }
            Ext.getCmp("popFoldCombo").folId = null;
            Ext.getCmp("popFoldCombo").clearValue("");
            var tid= Ext.getCmp("popolid").getValue();
            var grecno= Ext.getCmp("grecno").getValue();
            var gx = getActiveGrid();
            var rec = gx.getStore().getAt(grecno);
            var dpass = new Array();
            try{var xlist = [tid, rec, false];
            dpass.push(xlist);
            var xacts = {task:null,reply:null,appt:null, fwd: null, del: null, copy: null, cats: ""};
            moveOrCopyMessage(fdest, dpass,  xacts);
            }
            catch(e){}
        }}},
        {
        icon:'img\\folder.gif',
        cls:'x-btn-icon',
        tooltip: txtGoToFolder,
		hidden : ((it.Class == 43  && (isIgrid))?false:true),
		listeners:{click: function(b,e){
            var fdest = Ext.getCmp("popFoldCombo").folId;
            if( fdest == "" || fdest == null){
                (function(){ Ext.getCmp('popFoldCombo').focus();}).defer(100);
                return;
            }
            Ext.getCmp("popFoldCombo").folId = null;
            Ext.getCmp("popFoldCombo").clearValue("");
            outlookView(fdest, true);
        }}},
        {
			icon:'img\\reply.gif',
            cls:'x-btn-icon',
            tooltip: txtReply,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					try{var tid= Ext.getCmp("popolid").getValue();var it=NSpace.GetItemFromID(tid); replyToMessage(it);}catch(e){}
                }
			}
		},{
			icon:'img\\forward.gif',
            cls:'x-btn-icon',
            tooltip: txtFwd,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					try{var tid= Ext.getCmp("popolid").getValue();var it=NSpace.GetItemFromID(tid); var nit=it.Forward;nit.Display();}catch(e){}
                }
			}
		},{
			icon:'img\\copy.gif',
            cls:'x-btn-icon',
            tooltip: txtCopy,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					try{selFolder(txtCopy,popCopyMessage, null, false);}catch(e){}
                }
			}
		},{
			icon:'img\\task.gif',
            cls:'x-btn-icon',
            tooltip: txtNewTaskHere2,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
                click: function(b,e){
                try{var tid= Ext.getCmp("popolid").getValue();var it=NSpace.GetItemFromID(tid);
                var tsk = new Array();
                var catsWereEmpty=false;
				if( it.Categories == ""){
					it.Categories = "!Next";
                    catsWereEmpty = true;
                }
                tsk = archiveItem(it,null,null,true,false,true);
                if( catsWereEmpty){
                    it.Categories = "";
                    it.Save();
                }
                if( tsk.length == 1 && tsk[0].status == true){
                    var xarray = new Array();
                    var xao = new actionObject(tsk[0].task);
                    var xar = new actionRecord(xao);
                    xarray.push(xar);
                    editAction(xarray, true,false);
                }
                }catch(e){};
                }
            }
		},{
			icon:'img\\appoint.gif',
            cls:'x-btn-icon',
            tooltip: txtCrAppForInfo,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					try{var tid= Ext.getCmp("popolid").getValue();var it=NSpace.GetItemFromID(tid);
                var aptt=calendarItems.Items.Add();
				aptt.Subject=it.Subject;
				try{aptt.Body=it.Body;}catch(e){}
				aptt.Display();}catch(e){}
                }
			}
		},{
			cls:'x-btn-icon',
            icon: 'img\\info.gif',
            tooltip: txtOpenInOutlook,
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					try{var tid= Ext.getCmp("popolid").getValue();var it=NSpace.GetItemFromID(tid); it.Display(); }catch(e){}
                }
			}
		},{
            icon: 'img\\page_delete.gif',
            cls:'x-btn-icon',
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
            tooltip: txtDelete+' (DEL)',
            listeners:{click: function(b,e){
                var t=confirm(txtMsgDelItem);
                if (t==true)
                {
                var tid= Ext.getCmp("popolid").getValue();
                OLDeleteItem(tid);
                var grecno= Ext.getCmp("grecno").getValue();
                var gx = getActiveGrid();
                var rec = gx.getStore().getAt(grecno);
                // body of our display updated by updateRecordItem
                // which calls setBodyPopDisplay
                var idx= updateRecordItem(rec,true);
                }
                }
                }
        },{
			text: "Toggle View",
			hidden : ((it.Class == 43  && (isIgrid))?false:true),
			listeners:{
				click: function(b,e){
					var g = Ext.getCmp("bodytxt");
					var tid= Ext.getCmp("popolid").getValue();
					var it=NSpace.GetItemFromID(tid);
					if( getActiveGrid().getId() != "igrid")
						return;
	 				var x = Ext.getCmp("bodyval");

					if(g.viewval == "text")	 {
						x.body.dom.innerHTML=it.HTMLBody;g.viewval = "html";

					}else{
						x.body.dom.innerHTML=cleanBody(it.Body);g.viewval = "text";
					}

				}
			}
        }
    ];
    return tbar;
}
// move the displayed message to previous/next based on n
function goBodyPopMessage(n)
{
// get the grid and selection model
var grd=getActiveGrid();
var g=grd.getSelectionModel();
var psize=parseInt(jello.pageSize);
var bdisp = 0;
// get current rec displayed
var recno = parseInt(Ext.getCmp("grecno").getValue());
var idx = recno+n;
// if no selected item or requested item is start of first page of messages
if (g.hasSelection()==false || (idx <= 0 && lPage <= 1)){g.selectFirstRow();}
// is index less than 0 and not page 1 go to previous page, display is last item
else if( idx < 0){
    addToStore(null,null,(lPage +n), true,true, false);
    bdisp = psize-1;
}else if (idx >= psize){
// equal to psize means next page, dosplay is 0
   addToStore(null,null,(lPage +n), true,true, false);
}else{
// on the page
bdisp = idx;
}
setBodyPopDisplay(bdisp);

}
// using selected item display the body in the box
function setBodyPopDisplay(idx)
{
var gx = getActiveGrid();
if(idx == null ||  gx == "actgrid")
	return;
var init_t,init_type;
// get the record number - if no component than body not displayed
var reccmp = Ext.getCmp("grecno");
if(reccmp == null || typeof(reccmp) == "undefined")
    return;
//var gx = getActiveGrid();
var sm = gx.getSelectionModel();
// get the record
var rec = gx.getStore().getAt(idx);
var rid= rec.get("entryID");
reccmp.setRawValue(idx);
Ext.getCmp("popolid").setRawValue(rid);

try { var it = NSpace.GetItemFromID(rid);

if(getActiveGrid().getId() == "igrid"){
    // if popped from a folder display
    var x = Ext.getCmp("bodyval");
    updateMailItemDisplay(rec);
	mailTpl.overwrite(x.body, rec.data);
}else{
    // task display
    var pnl = Ext.getCmp("bpanel");
    pnl.setValue(it.Body);
    pnl.render();
}
// set title to have folder and n/m
var ntitle= it.Parent.FolderPath + "(" + (idx+1).toString() +"/" + it.Parent.Items.Count +")";
var g = Ext.getCmp("bodytxt");
var pfd = Ext.getCmp('pformtagdisplay');
if( typeof(pfd) != "undefined" && pfd != null){
pfd.getEl().dom.innerHTML="";
Ext.getCmp('ptagfield').setValue(it.Categories);
var f=tagList('pformtagdisplay',null,Ext.getCmp('ptagfield').getValue(), rid);
var s=f[0];
pfd.getEl().dom.innerHTML=s;
pfd.render();
Ext.getCmp('ptags').reset();
}
g.setTitle(ntitle);
if( sm.isSelected(idx))
    return;
sm.clearSelections();
sm.selectRow(idx);
}
catch(e){}
}
// set category for selected item and update display
function setPopCategory()
{
    var id = Ext.getCmp('popolid').getValue();
    try{
        var it = NSpace.GetItemFromID(id);
        it.Categories = Ext.getCmp('ptagfield').getValue();
        it.Save();
        var recno = parseInt(Ext.getCmp("grecno").getValue());
        updateInboxRecord(recno);
    }catch(e){}
}
//copy message for item displayed in opup
function popCopyMessage(dest)
{
var init_t,init_type;
var recno = parseInt(Ext.getCmp("grecno").getValue());
var gx = getActiveGrid();
var rec = gx.getStore().getAt(recno);
var rid= rec.get("entryID");
try{
var t = NSpace.GetFolderFromID(dest);
var  it = NSpace.GetItemFromID(rid);
if( t != null)copyMailTo(it,t);
}catch(e){}

}

var lastView = null;
function setActionView()
{
 	var g = getActiveGrid();
 	if( g.getId() == 'grid')
		var iF=NSpace.GetFolderFromID(jello.actionFolder);
	else{
		var fldx = lIbFolder;
		if( fldx == "" || fldx == null)
			ctx = jello.inboxFolder;
		var iF=NSpace.GetFolderFromID(ctx);
	}
 	var views = iF.Views;
 	if(views.Count == 0)
 		return;
 	var varray = new Array();
 	for( var i=1; i <= views.Count; i++)
 	{
 		try{var vw = views.Item(i);
 		varray.push(new Array(vw.name, vw.Filter));
 		}catch(e){continue;}
 	}
 	vnameStore = new Ext.data.SimpleStore({
				fields: ['name','filter'],
				data : varray
			});
	var combo = new Ext.form.ComboBox({
			id: "selActionView",
			store: vnameStore,
			anyMatch: true,
			matchWordStart: false,
			queryIgnoreCase: true,
			displayField:'name',
			typeAhead: true,
			mode: 'local',
			//emptyText:txtGotoInfo+'...',
			selectOnFocus:true,
			resizable: true,
			minListWidth: 110,
			listeners : {
				select : function(combo,record,idx){

					var iF=NSpace.GetFolderFromID(jello.actionFolder);
					var g = getActiveGrid();
					var ctx = g.ctx;
					if( g.getId() == "grid"){
					var flt= record.get("filter");
					flt = (flt==""?"":"("+flt+") AND ") +
						"(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "')";
          			showContext(ctx,1,flt,4);
          			}else{
          				ctx = lIbFolder;
						if( ctx == "" || ctx == null)
          					ctx = jello.inboxFolder;
          				var iF=NSpace.GetFolderFromID(ctx);
          				var flt = record.get("name");
          				var vws = iF.Views;
          				var v = vws.Item(flt);
          				v.Apply();
          				pInbox(lIbFolder);
          			}
					setTimeout(function(){
        			var w = Ext.getCmp("selectview");
							w.close();},6);
				}
			}
		});
	var win = new Ext.Window({
        title: "Select View",
        width: 300,
        height: 100,
        id:'selectview',
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: combo,
        modal:true,
        listeners: {destroy: function(t){showOLViewCtl(true);}},
        buttons: [
			{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();}
            }}

        ]
    });

    showOLViewCtl(false);
	win.show();

    win.setActive(true);
  setTimeout(function(){Ext.getCmp("selActionView").focus(true,100);},1);

	//shortcut keys
var map = new Ext.KeyMap('selActionView', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});
 }
 function setDueDate(it,due)
 {
 	try{
	 	var dt=new Date(due);
		//due=DisplayDate(dt);
		due=dt.format("Y/m/d");
	 	if (it.Class!=48)
			{//non task
			var dp=getJDueProperty(it);
			setJDueProperty(it,due);
			it.Save();
			}
		else
			{//task
			it.itemProperties.item(dueProperty).Value=dt.format("Y/m/d");
			it.Save();
			}
	}catch(e){return true;}
	return false;

 }

 function makeNextActionStep(id)
 {
 	var it = NSpace.GetItemFromID(id);
 	if( it.Complete == true)
 		return;
 	var dt = new Date();
 	it.StartDate = dt.format("Y/m/d");

 	// found a setp indicator
 	var subj = it.Subject;
 	var sregex = new RegExp("^"+txtStep + " \\d+: ");
	var step = subj.match(sregex);
 	var oldstep = txtCompleted + (step==null?" " + txtStep +" 1":" "+step.toString().replace(":","")) + " @ " + dt.format(jelloDateFormatString()) + " : " + subj.replace(sregex,"") +	"\n";

	var uStep = getStepFromBody(it,oldstep);
	// null means empty section, ask about completing
	if( uStep == null){
		var ans = confirm(txtCompleteItem);
		if( ans != true)
			return;
		var notes = getOLBodySection(it,jelloNotesDelim);
		notes = notes + "\n-*** "+ txtCompleted + " @ " + dt.format(jelloDateFormatString()) + " ***\n";
		setOLBodySection(it,jelloNotesDelim,notes);
		it.Save();
		actionDone(id);
		return;
	}

	dt.setDate(dt.getDate()+uStep[1]);
 	it.DueDate = dt.format("Y/m/d");
	// use same subj?

	if( step !=  null)
	 {
	 	var snumb = parseInt(step.toString().match(/\d+/).toString());
	 	snumb++;
	 	if( uStep[0] == "")
	 		subj = subj.replace(sregex, txtStep + " " + snumb +": ");
	 	else
	 		subj = txtStep + " " + snumb +": " + uStep[0];
	 }else{
		if( uStep[0] == "")
		 		subj = txtStep+" 2: " + subj;
	 	else
	 		subj =txtStep+" 2: " + uStep[0];
	}
	 it.Subject = subj;
	 it.Save();

 }

 function getStepFromBody(it,oldstep)
 {
 	// get the steps section
	var notes = getOLBodySection(it,jelloNotesDelim);
	if( notes != null && notes.charAt(notes.length-1) != '\n')
		notes += '\n';
	var steps =  getOLBodySection(it,jelloStepsDelim);
	if( steps != null)
		steps = steps.replace(/^ *$/g,"");
	var subject = new Array("");
	var i = null;
	var dueOffset = 7;
	var comments = "";
	if( steps != "" &&  steps != null)
	{
		// split into lines
		steps = steps.replace(/\r/g,"");
		steps = steps.toString();
	 	var stp = steps.split("\n");
	 	// top line is next step
	 	// but might be comment lines
	 	// if so, move to notes



		 for( i=0; i < stp.length; i++)
	 	{
	 		var theline = stp[i];
	 		if( theline == "")
	 			continue;
			 // comment line?  if so add to top of notes
	 		if(theline.indexOf("//") == 0){
	 			comments = comments +  theline + "\n";
	 			continue;
	 		}
			subject = theline.split(",");
			// if 2 args and second is only digita, it is an offset
			if( subject.length == 2  && subject[1].trim().match(/\D+/) == null){
				dueOffset = parseInt(subject[1].trim(	));
				if( dueOffset < 0)
					dueOffset = 7;
				break;
			}

	 	}
	}
	// oldstep ends in a newline
	setOLBodySection(it,jelloNotesDelim,notes + oldstep  + comments);
	// if no steps found - retun null menaing complete
	// no need to update steps since they are emoty
	if( steps == "" )
		return null;
	// steps niull means no steps section
	// i = length meant we processed comments or empty lines
	// but no step found
	// in either case return empty strings so we get default step
	// subject updates
	if( steps == null || i == stp.length){
		subject[0] = "";
		steps = "";
	}else{
		stp = stp.slice(i+1);
		steps = stp.join("\n");
	}
	setOLBodySection(it,jelloStepsDelim,steps);
	return (new Array(subject[0],dueOffset));

 }
function editStep(id,rec)
{
	var it = NSpace.GetItemFromID(id);
	var steps = getOLBodySection(it,jelloStepsDelim);

	var tarea= new Ext.form.TextArea({
        	fontsize: jello.mailPreviewFontSize,
        	width: 280,
        	id: "steptext",
        	value: steps
        });
	var win = new Ext.Window({
        title: txtCreateSteps,
        width: 480,
        height:370,
        id:'stepwin',
        minWidth: 300,
        minHeight: 280,
        resizable:true,
        draggable:true,
        layout: 'fit',
        plain:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: tarea,
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text: txtSave,
            listeners:{
            click: function(b,e){
	       		var steps=Ext.getCmp("steptext").getRawValue();
				var id = Ext.getCmp("stepwin").olid;
				var it = NSpace.GetItemFromID(id);
				setOLBodySection(it,jelloStepsDelim,steps);
				it.Save();
				var g=getActiveGrid();
				var store = g.getStore();
				var recno = store.find("entryID",id);
				if( recno != -1){
					rec = store.getAt(recno);
					rec.set("body", getItemBody(it));
				}
				g.getView().refresh();
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
        }],
        olid: id
    });

    showOLViewCtl(false);
    win.show();
    win.setActive(true);

}

function convertToProject(id)
{
//convert action to project
if (typeof(id)=="number")
{
var r=getTagByID(id);
		if (r!=null)
		{
		var pj=r.get("isproject");
		r.beginEdit();
		r.set("isproject",!pj);
		r.endEdit();
		syncStore(tagStore,"jello.tags");
		jese.saveCurrent();

		}
return;
}
try{var it=NSpace.GetItemFromId(id);}catch(e){alert(txtInvalid);return true;}
createTag(it.Subject,false,lastOpenTagID,false,false,true);
it.Delete();
}

function actionDateEdited(obj)
{

	var store = obj.grid.getStore();
	try{ var it = NSpace.GetItemFromID(obj.record.get("entryID"));}catch(e){return;}

  if (obj.field=="subject")
  {
  it.Subject=obj.value;
  it.Save();
	obj.record.commit();
	store.commitChanges();
  return;
  }
	
  // if a tickler grid, recalc the daypos for grouping
	setDueDate(it,obj.value);
	if( obj.grid.getId() == "tgrid")
	{
		var dt = new Date(obj.value);
		obj.record.set("start",dt);
		obj.record.set("daypos",getDayPosition(dt));
	}

	it.Save();
	obj.record.commit();
	store.commitChanges();
	if(obj.grid.getId() == "tgrid")
		store.groupBy(store.groupField,true);
	var idx=columnStore.find("grid",new RegExp("^"+obj.grid.getId()+"$"));
	if (idx==-1){return;}
	var r=columnStore.getAt(idx);
	var newConfig=Ext.decode(r.get("data"));
	var doSortFld=null;var doSortOrd=null;
	var cm = obj.grid.getColumnModel();
	for (var i = 0; i < newConfig.length; i++) {
	  var ncdi=newConfig[i].dataIndex;
	  try{
	  if (newConfig[i].sorted!=null){doSortFld=newConfig[i].dataIndex;doSortOrd=newConfig[i].sorted;break;}
	  }catch(e){}
	}

	try{store.sort(doSortFld,doSortOrd);}catch(e){}
    obj.grid.render();

}

//create a status report for a task or series of tasks

function emailActionItem(cit)
{
//send status report
    var xe = /<br>/g;
    var le = /\r/g;
    var rmn = /\n\n\n+/g;
    var msg = NSpace.GetFolderFromID(jello.inboxFolder).Items.Add();
    var empty = true;
    for( var i=0; i < cit.length; i++)
    {
        var srpt = NSpace.GetItemFromID(cit[i].get("entryID")).StatusReport();

        var s = srpt.Body.replace(xe,"\n");
        s = s.replace(le,"");
        s = s.replace(rmn,"\n\n");
        msg.Body += s;
        if( srpt.To != "" & srpt.To != null)
            msg.To += ";"+srpt.To;
        if( srpt.CC != "" & srpt.CC != null)
            msg.CC += ";"+srpt.CC;
        srpt.Delete();

    }
    msg.Display();
 }

 

function processSentItems()
{

	var ans= prompt("Enter period in days for processing Sent items\n(0 means since midnight, blank to skip)", "0");
	if(ans == "" || ans == null)
		return;
	var period = parseInt(ans);
	if( period < 0 )
		return;
	// we need a unique identifier to find the tasks we're about to create
	// use the utc string of the date and store in mileage, we'll clear later
	var xd = new Date();
	var tagger = xd.toUTCString();
	// but we'll also figure out the dasl to get only iyems for the specified period
	xd.setHours(0,0,0,0);
	xd.setDate(xd.getDate()-period);
	var tdasl = "( urn:schemas:httpmail:datereceived >= '"+ xd.format(jelloDateFormatString(true)) + "')";
	// get the name of the default folder, likely the display name for the ser
	var nm = NSpace.GetFolderFromID(jello.inboxFolder).Parent.FolderPath;
	var x1 = nm.indexOf("@");
	// claen it up
	nm = nm.substring(0,x1);
	nm = nm.replace(/_/g," ");
	nm = nm.replace(/\\/g,"");
	var debug = false;
	// build list of matches for bcc search
	var emailaddr= nm+";"+jello.bccMatch;
	// build dasl
	var dasl = "@SQL=((http://schemas.microsoft.com/exchange/mileage IS NULL) AND (";
	var addlist = emailaddr.split(";");
	for( var i=0; i < addlist.length; i++){
		if( addlist[i] == "")
			continue;
		if( i != 0)
			dasl += " OR ";
		dasl += "(urn:schemas:calendar:resources LIKE '%" +addlist[i]+"%')";
	}
	dasl += ") AND " + tdasl + ")";
	// filter sent
    var its = sentItems.Items.Restrict(dasl);
	    if( its.Count == 0 || debug){
        Ext.Msg.alert("","No items found");
		return;
    }
	// now get the tag to apply
	var theTag = (jello.bccCategory == ""?"@Waiting":jello.bccCategory);
	// have items, now create tasks
	for( var i=1; i <= its.Count; i++){
		var it = its.Item(i);

		var ta = new Array();
		var cats = it.Categories;
		try{if(it.Categories != "")
			it.Categories += ";";
		it.Categories +=  theTag;
		// put tag in Mileage so we know not to process again
		it.Mileage = theTag;
		it.Save();
		}catch(e){}
		var res = archiveItem(it,null,false,true,false,true)[0];
		// set mileage so we can fild the ones we just created
		//it.Categories = cats;
		//it.Save();
	}
	//var titms = NSpace.GetFolderFromID(jello.actionFolder).Items.Restrict("@SQL=http://schemas.microsoft.com/exchange/mileage = '"+
	//	tagger+"'");
	//for( var i=1; i <= titms.Count; i++)
	//{
	//	var xi = titms.Item(i);
	//	xi.Mileage = "";
	//	xi.Save();
	//}
	// now show the grid to edit
	var xtag = theTag.split(";");
	var tg = tagStore.find("tag",Ext.escapeRe(xtag[0]));
	var r = tagStore.getAt(tg);
	var tid = r.get("id");
	showContext(tid,1);
    //buildGrid(null,titms,0,0,null,true,false);

}

function toggleEditableGrid()
{
 var ised=Ext.getCmp("aeditable").pressed;
 var grd=Ext.getCmp("grid");
  if (ised==false)
  {
  //disable in cell editing
  grd.getColumnModel().setEditable(5,false);
  }
  else
  {
  //enable in cell editing 
  grd.getColumnModel().setEditable(5,true);
  
  }
}

function copyActionsToTag(id,editems)
{
//copy selected actions to another tag
   var simple = new Ext.FormPanel({
        labelWidth: 75, // label settings here cascade unless overridden
        frame:true,
        title: 'Copy actions to another Tag',
        width:480,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        iconCls:'actionformicon',
        id:'actcopyform',
        buttonAlign:'center',
        defaults: {width: 300},
        defaultType: 'textfield',

        items: [
                new Ext.form.ComboBox({
                fieldLabel: 'Select or Create Tag',
                id: 'copytags',
                store:globalTags,
                hideTrigger:false,
                typeAhead:true,
                mode:'local'
            }),            
            new Ext.form.Label({
				    html:editems.length+' Item(s) to Copy',
				    id:'actioncopyitems',
				    height:90,
				    cls:'formattlist'
            })
            ]
});

   var win = new Ext.Window({
        title: 'Copy to tag',
        width: 500,
        height:200,
        id:'theactioncopyform',
        minWidth: 300,
        minHeight: 200,
        resizable:false,
        draggable:false,
        layout: 'fit',
        plain:true,
        draggable:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        modal:true,
        items: simple,
        listeners: {destroy:function(t){try{showOLViewCtl(true);}catch(e){}}},
        buttons: [{
            text:'<b>'+txtCopy+'</b>',
            id:'sbut',
            tooltip:txtCopy,
            listeners:{
            click: function(b,e){
            copyTheActions(b,editems);
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

//show the form
    showOLViewCtl(false);
    win.show();
    win.setActive(true);
  setTimeout(function(){try{Ext.getCmp("copytags").focus(true,100);}catch(e){}},1);

  var actionKeyMap = new Ext.KeyMap('theactioncopyform',
	[
	{

    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
            try{win.destroy();}catch(e){}

            },
    scope: this}

    ]);

}

function copyTheActions(b,recs)
{

var toTag=Ext.getCmp("copytags").getValue();
  var nTag=getTag(toTag);
    if (nTag==null)
    {
    //Add not existing tag
    var ntt=createTag(toTag,false,0,false,true,false);
    if (ntt==false){alert("This tag cannot be added. Aborting!");Ext.getCmp("theactioncopyform").destroy();return;}
    else{alert("Tag "+toTag+" created");}
    }
var noRec=false;
var counter=0;
	for (var x=0;x<recs.length;x++)
	{
		var r=recs[x];

		if (typeof(r)=="object")
		{
		try{
		var id=r.get("entryID");}catch(e)
			{
			var id=r.EntryID;noRec=true;
			}
		}
		else
		{var id=r;}

    var it=NSpace.GetItemFromID(id);
    try{
      var newIt=it.Copy();
      newIt.itemProperties.item(catProperty).Value=toTag;
      newIt.Save();
      counter++;
       }catch(e){alert("Could not copy the Item "+it.Subject);}
   }
   alert(counter+" Item(s) copied to "+toTag);
   Ext.getCmp("theactioncopyform").destroy();
}
