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

var masterExpanded=true;

 var aReader = Ext.data.ArrayReader([
    {name: 'subject'},
    {name: 'type'},
    {name: 'entryID'},
	{name: 'importance'},
	{name: 'attachment'},
	{name: 'created',type:'date'},
	{name: 'iclass'},
	{name: 'contacts'},
	{name: 'sensitivity'},
	{name: 'sortDate'},
	{name: 'status'},
	{name: 'due',type:'date'},
	{name: 'icon'},
	{name: 'unread',type:'boolean'},
	{name: 'tags'},
	{name: 'tag'}


]);


function pMaster(filter)
{
//masterlist rendering
masterExpanded=true;
lastContext="";
thisGrid="mgrid";
initScreen(true,"pMaster('"+filter+"')");
getTagsArray();
 //---get items----
var listStore=getMasterItems(filter);
var counter=listStore.getCount();
//-------------------- 

var g=getActiveGrid();
try{g.destroy();}catch(e){}

//create grid
var gridwaitMask = new Ext.LoadMask(Ext.getBody(), {msg:txtMsgWait+"..."},{store:listStore});
sm= new Ext.grid.CheckboxSelectionModel({});

/*
var tfoot = new Ext.Toolbar({
 id:'gridfooter',
items:[
                new Ext.form.ComboBox({
                id: 'filter',
                store:filterValues,
                hideTrigger:false,
                valueField:'value',
                displayField:'text',
                triggerAction:'all',
                selectOnFocus:true,
				editable:false,
                emptyText:txtShow+'...',
                mode:'local',
                hidden:false,
                listeners:{
                specialkey: function(f, e){
				if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                filterActions();}},
				select: function(cb,rec,idx){
	            filterMaster();
		        }
                }
                })

	] });
 */
var dts=jello.dateSeperator;
var fmt="n"+dts+"j"+dts+"y";
var mtbar=addActionToolbar(true);
var expbtn=new Ext.Button({
	tooltip: 'Collapse All',
	cls:'x-btn-icon',
    id:'aexpand',
    icon: 'img\\collapse.gif',
    handler:expandCollapseAll
});
mtbar.insert(0,expbtn);
var ALinline=false;if(jello.ALinlineduedate==1 || jello.ALinlineduedate=="1"){ALinline=true;}

//var detailPanel = Ext.getCmp('previewpanel');
//detailPanel.expand(true);
    var grid = new Ext.grid.EditorGridPanel({
        store: listStore,
        id:'mgrid',
        tbar:mtbar,
        bbar:mtbar,
		clicksToEdit: 1,
        columns: [
        sm,
			{header: "", width: 20, fixed:true, sortable: false, renderer: getImportance, dataIndex: 'importance'},
			{header: "", width: 30, hideable:false, fixed:true, sortable: false, renderer: getIcon, dataIndex: 'icon'},
            {header: "", width: 20, hidden:true, fixed:true, sortable: false, dataIndex: 'sensitivity'},
            {header: txtSubject, width: 350, sortable: true, renderer: renderMSubject,dataIndex: 'subject'},
            {header: txtDueDate, width: 80, sortable: true,hidden:false, renderer:DisplayAppDate, dataIndex: 'due',
			 editable: ALinline,editor: new Ext.form.DateField({
                    format: fmt,
                    editable: false
                })
             },
            {header: txtCreated, width: 80, hidden:true, sortable: true, renderer:DisplayAppDate, dataIndex: 'created'},
            {header: txtNotes, width: 100, hidden:true, sortable: true, renderer: renderMNotes, dataIndex: 'body'},
            {header: txtTag, width: 100,hidden:true, sortable: true, dataIndex: 'tag'}
        ],
        stripeRows: true,
        autoScroll:true,
        deferRowRender:false,
        region:'north',
        enableColumnHide:true,
           view: new Ext.grid.GroupingView({
            forceFit:true,
            groupTextTpl: '{text} ({[values.rs.length]} {[values.rs.length > 1 ? "Items" : "Item"]})'
        }),

        viewConfig:{
        emptyText:txtNoDispItms
        },
        trackMouseOver:true,
//height:Ext.getBody().getViewSize().height-jello.actionPreviewHeight,
height:200,
		sm:sm,
    listeners:{
    mouseover: function(e){thisGrid='mgrid';},
		rowdblclick: function(g,row,e){
		openMasterItem();
        },		
        cellcontextmenu: function(g, row,cell,e){
			rightClickItemMenu(e,row,g);
		 },
		 afteredit: function (obj){
                actionDateEdited(obj);
		 }

		},
        enableColumnMove:true,
        border:true,
        padding:0,
        
        //layout:'fit',
        //width:panelWidth-5,
        loadMask:gridwaitMask
  
        
    });

    //main.innerHTML="<div id=toolbar></div><div id=list></div>";
    //grid.render('list');
    //addActionToolbar(true);
    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.doLayout();

//*150410*//
	 	grid.getSelectionModel().on('rowselect', function(sm, rowIdx, r) {
		var detailPanel = Ext.getCmp('previewpanel');
		actionTpl.overwrite(detailPanel.body, r.data);
				//var gb=r.get("body");
	   });
//************	   
	   
    grid.on('columnresize',function(index,size){saveGridState("mgrid");});
    grid.on('sortchange',function(){saveGridState("mgrid");});
    grid.getColumnModel().on('columnmoved',function(){saveGridState("mgrid");});
    grid.getColumnModel().on('hiddenchange',function(){saveGridState("mgrid");});
    restoreGridState("mgrid");

ppnl.setTitle("<img src='img//master16.png' style=float:left;>&nbsp;"+txtMaster+getFilterName(filter));
ppnl.setAutoScroll(false);
setTimeout(function(){
try{
if (jello.selectFirstItem==1 || jello.selectFirstItem=="1"){grid.getSelectionModel().selectRow(0);}
grid.getView().focusRow(0);grid.focus();}catch(e){}
resizeGrids();
},16);


// map one key by key code
var mmap = new Ext.KeyMap("mgrid", {
    key: 13,
    fn: function(){openMasterItem();},
    scope: this
});

//ctrl + INS adds action
var mmap1 = new Ext.KeyMap("mgrid", {
    key: Ext.EventObject.INSERT,
    fn: function(){editAction(null);},
    ctrl:false,
    shift:false,
    stopEvent:true,
    scope: this
});

//DEL deletes items
var mmap2 = new Ext.KeyMap('mgrid', {
    key: Ext.EventObject.DELETE,
    fn: function(){
    actionSelected(Ext.getCmp('adel'));
    },
    stopEvent:true,
    ctrl:false,
    scope: this
});

//ctrl+q open in outlook
var mmap2 = new Ext.KeyMap('mgrid', {
    key: 'q',
    ctrl:true,
    fn: function(){
    actionSelected(Ext.getCmp('aoutlook'));
    },
    stopEvent:true,
    ctrl:false,
    scope: this
});
//status=txtReady;
updateCounter(counter);
//try{Ext.getCmp("gridfooter").setHeight(25);}catch(e){}
 }

function getMasterItems(filter)
{
//get master list items 

var store= new Ext.data.GroupingStore({
reader: reader,
sortInfo:{field: 'tag', direction: "ASC"},
id:'masterStore',
groupField:'tag',
        fields: [
    {name: 'subject'},
    {name: 'type'},
    {name: 'entryID'},
	{name: 'importance'},
	{name: 'attachment'},
	{name: 'created',type:'date'},
	{name: 'iclass'},
	{name: 'contacts'},
	{name: 'sensitivity'},
	{name: 'sortDate'},
	{name: 'status'},
	{name: 'due',type:'date'},
	{name: 'icon'},
	{name: 'unread',type:'boolean'},
	{name: 'tags'},
	{name: 'tag'}
        ]
    });
    
reviewStore=tagStore;

//reviewStore.filter("istag",true);
//reviewStore.filter("isprivate",false);
reviewStore.filter([
{property:'istag',value:true}
//,{property:'private',value:!true}
]);
reviewStore.sort("tag","ASC");
var counter=0;

var masterTag = function(r){
var tname=r.get("tag");
var tid=r.get("id");
var tarc=r.get("archived");
if (tarc){return;}
var tpri=r.get("private");
if (tpri){return;}
try{var its=showContext(tname,0,null,filter,null,true,true);}catch(e){}
//status=txtStatusGetting+"..." + tname;
//its.Sort("http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040",true);
var olc=its.Count;
if(olc>0)
{
    for (var x=1;x<=olc;x++)
    {
    				var ar=new actionObject(its(x));
    				ar.tag=tname;
					var newRec=new actionRecord(ar);
    				store.add(newRec);
    				counter++;
    				status=txtStatusGetting+" ["+tname+"]..."+x+"/"+olc;
		}
}
		
};
reviewStore.each(masterTag);
store.sort("subject","ASC");
status="";
reviewStore.clearFilter();
return store;
}

function renderMSubject(v,m,r)
{
//render the subject line in grid
var id=r.get("entryID");
var alltag=r.get("tags");
var tt=r.get("tag");
var st=r.get("status");
var cl=r.get("iclass");
var un=r.get("unread");
var con=r.get("contacts");
var rt="";
var nas=getActionTagName();
if (tt==txtNew){var tList=tagList("formtagdisplay","...",alltag);}
  var subj=Ext.util.Format.htmlEncode(v);
  if (st==2){subj="<strike>"+subj+"</strike>";}
  if (un==true){subj="<b>"+subj+"</b>";}
  var conicon="";
  var con2=con.replace(new RegExp("'","g")," ");
  if (notEmpty(con)){conicon="<a class=jellolink title='"+con2+"'><img src="+imgPath+"user.gif></a>";}
  rt+=NextActionToggleIcons(alltag,id)+conicon+"&nbsp;"+subj+"&nbsp;";
  if (tt==txtNew){rt+="&nbsp;"+tList[0];}
  else{
        alltag=alltag.replace(new RegExp(";","g")," ");
        alltag=alltag.replace(new RegExp(nas,"g"),"");
      rt+="&nbsp;<span class=caltime>"+alltag+"</span>";
      }
 return rt;
}

function openMasterItem()
{//open inbox item for editing
var g=Ext.getCmp("mgrid").getSelectionModel();
var r=g.getSelected();
var id=r.get("entryID");
var it=NSpace.GetItemFromID(id);
if (it.Class==48){var t=new Array(r);editAction(t);}
else{it.Display();}
}

/*
function filterMaster()
{
//filter masterlist
var f=Ext.getCmp("filter").getValue();
pMaster(f);
}
*/

function getFilterName(filter)
{
//return filter name
if (!notEmpty(filter)){return "";}
var ix=filterValues.find("value",filter);
var r=filterValues.getAt(ix);
return " <i>["+r.get("text")+"]</i>";
}

function printMasterList(direct)
{
//preview & print masterlist
var store=Ext.getCmp("mgrid").getStore();
var counter=0;
var lastTag="";
var nas=getActionTagName();
var ret="<head><style>.printChars{height:30px;font-size: 12px;font-family: 'Segoe UI', Verdana, 'Trebuchet MS', Arial, Sans-serif;border-bottom: 1px solid Gainsboro;}.smallPchars{font-size:9px;color:black;}</style></head><body><table width=100% cellpadding=0 cellspacing=0>";
ret+="<tr><td class=printChars><b>"+txtSubject+"</b></td><td class=printChars><b>"+txtDueDate+"</b></td></tr>";

  var toPrint = function(rec){
  try{var ddd=DisplayDate(rec.get("due"));}catch(e){var ddd="";}
  var tg=rec.get("tag");
  if (lastTag!=tg){lastTag=tg;ret+="<tr><td colspan=2 valign=bottom class=printChars style=background:#FAEFC7;height:45px;><b>"+tg+"</b></td></tr>";}
  var alltag=rec.get("tags").replace(new RegExp(tg,"g"),"");
  alltag=alltag.replace(new RegExp(";","g")," ");
  alltag=alltag.replace(new RegExp(nas,"g"),"");
  var content=rec.get("subject");
  if (alltag.length>1){content+="&nbsp;<i>("+alltag+")</i></span>";}
  
  var nts=rec.get("body");
  var s=nts.search(txtTryToGetBody);
  if(notEmpty(nts) && s==-1){content+="<br><span class=smallPchars>"+nts.substr(0,500)+"</span>";}
  else
  {content+="<br><span class=smallPchars>"+rec.get("notes").substr(0,500)+"</span>";}
  ret+="<tr><td class=printChars><span style=font-family:Wingdings;font-size:16px;>q</span> "+content+"</td><td class=printChars>"+ddd+"</td></tr>";
  counter++;
  };

store.each(toPrint);  
ret+="</table><br><br><span class=printChars>"+counter+" "+txtItemItems+" @ "+DisplayDate(new Date())+"</span>";


var ff=NSpace.GetFolderFromID(jello.inboxFolder).Items;
var it=ff.Add(6);
//lastOpenTagID
try{it.Subject=txtMaster;}catch(e){it.Subject=txtMaster;}
it.HTMLBody=ret;
if (direct==true){it.PrintOut();Ext.info.msg(txtPrinting,txtPrintOK);}
else{it.Display();}

}


function renderMNotes(v,m,r)
{
var ret="";
var s=v.search(txtTryToGetBody);
if (s>-1){ret= "";}else{ret=v;}
var n=r.get("notes");
return ret+" "+n;
}

function expandCollapseAll()
{
 if (masterExpanded)
 {
 var bt=Ext.getCmp("aexpand").setIcon("img\\expand.gif");
 var g=Ext.getCmp("mgrid").getView().collapseAllGroups();
 bt.setTooltip("Expand All");
 masterExpanded=false;
 }
 else
 {
 var bt=Ext.getCmp("aexpand").setIcon("img\\collapse.gif");
 var g=Ext.getCmp("mgrid").getView().expandAllGroups();
 bt.setTooltip("Collapse All");
 masterExpanded=true;
 }
}
