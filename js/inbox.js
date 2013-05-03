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

var lStart=1;
var lEnd=0;
var lPage=0;
var lastInboxItems;
var lastListStore;
var isInbox=false;
var lIbFolder="";
var folderOnView=0;
var lastprevsz ="";
var lastArchivers="";
var otherArchivers="";
var refreshWholeView=false;
var lastInboxUpdateTime=null;
var inboxIntervalTimer=null;

// the next arrays select field ffor table depending on folder type, last maps to array obj
// order of first 2 items must be entryID and type of date and must always be this
// see code in outlook.js as the getTable has these always and the getFolderTable adds the others
var emailColumns = ['EntryID','ReceivedTime','Subject', 'Importance', 'SenderName', 'To', 'CC', 'urn:schemas:httpmail:textdescription',
	'urn:schemas:httpmail:hasattachment','TaskDueDate', 'UnRead',"MessageClass", "Sensitivity", "Categories"];
var taskColumns = ['EntryID','CreationTime','Subject','Importance', '', '', '', 'urn:schemas:httpmail:textdescription',
	'urn:schemas:httpmail:hasattachment','DueDate', 'UnRead',"MessageClass", "Sensitivity","Categories"];
var emailStoreObjNames = ['entryID','created','subject', 'importance', 'sender', 'to', 'cc', 'body', 'attachment',
	'due', 'unread','iclass', 'sensitivity','tags'];
// filed to sort on for various item types
var sortTables = ["ReceivedTime","Start","FullName",
	"DueDate","LastModificationTime","LastModificationTime","ReceivedTime", "CreationTime"];
// true for descending, false for ascending
var inboxSortOrder = [true,false,false,false,false,true,true,false];


function getInboxItems(cmd,istore,olFolder,dasl,page,clear)
{
//get inbox items and add them to a store
//cmd=0 count, cmd=1 add items to store
//if page=0 add all records, otherwise go to specific page using jello.pageSize
//clear=true clear store before adding new records
var theFolder = olFolder;
if( theFolder.Class != 2)
	theFolder = theFolder.Parent;
if( OLversion >= 12 && useTables && (theFolder.DefaultItemType == 0 || theFolder.DefaultItemType ==3)){return getInboxTableItems(cmd,istore,olFolder,dasl,page,clear);}


//build query
lPage=0;
var iits;
if (olFolder.Class==2)
{
var iF=olFolder.Items;
if (dasl!=null && dasl != "")
{iits=iF.Restrict("@SQL="+dasl);}
else
{iits=iF;}
iits.Sort(sortTables[olFolder.DefaultItemType],inboxSortOrder[olFolder.DefaultItemType]);
}else{iits=olFolder;}
var counter=iits.Count;
if( counter == 0)
	return counter;

//switch(olFolder.DefaultItemType){
//	case 1:
//		// appointments
//		iits.Sort("urn:schemas:calendar:dtstart",false);
//		break;
//	case 2:
//		// contacts
//		iits.Sort("urn:schemas:contacts:cn",false);
//		break;
//	case 3:
//		// tasks
//		iits.Sort("http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040",false);
//		break;
//	default:
//		iits.Sort("[urn:schemas:httpmail:datereceived]",true);
//		true;
//}


//add actions
if (cmd==1){addToStore(iits,istore,page,clear,false);}

return counter;
}

function addToStore(olitms,istore,page,clear,updateButtons)
{//add outlook item set in store
//alert(isInbox +" "+ page  +" "+ updateButtons);
// if olitems is not an Items collection then tables


status="Connecting and Calculating...";
try{var g=Ext.getCmp("igrid");g.hide();}catch(e){}
if (isInbox==true && page==1 && updateButtons==true){pInbox();return;}
if (olitms==null)
{olitms=lastInboxItems;istore=lastListStore;}
else if(typeof(Ext.getCmp("gotoStore")) == "undefined")
{lastInboxItems=olitms;lastListStore=istore;}
if( olitms.Class != 16 && OLversion >= 12 && useTables){return addTableItemsToStore(olitms,istore,page,clear,updateButtons);}

//clear store
if (clear==true){istore.removeAll();}


	var psize=0;
	psize=parseInt(jello.pageSize);
	var allCount=olitms.Count;

      if (page>-999)
      {
    
    	if( page <= 0)
    		page = 1;
    	if (page>0)
    	{	if( (page * psize) > (((allCount + psize -1) /psize)*psize))
    			page = parseInt((allCount + psize -1) /psize);
    		lStart=(((psize*page)-psize)+1);
    		lEnd=(lStart+psize)-1;}
      else{lStart=0;lEnd=(allCount> psize? psize:allCount);}
    	try{Ext.getCmp("inboxgotopage").setRawValue(""+page);}catch(e){}
    	if (lStart<1){lStart=1;}
    
      }
      else
      {
      //for inbox tasks there is no paging
      lStart=1;lEnd=allCount;
      }
    	
      if (lEnd>allCount){lEnd=allCount;}
    		if (olitms.Count>0)
    		{
    	
      var ifname=olitms(1).Parent.Name;
		  
			for (var x=lStart;x<=lEnd;x++)
			{
				var ar=new actionObject(olitms(x));
				var newRec=new actionRecord(ar);
				istore.add(newRec);
				status=txtStatusGetting+" "+ifname+" ..." + x+"/"+lEnd;
			}
	
		}
	 else{}

lPage=page;
if (isInbox==true){allCount+=inboxTasksCount();}


  if (updateButtons==true){updatePagingToolbar(istore,allCount);}
  try{g.show();g.getView().focusRow(0);}catch(e){}

//status=txtReady;
}


function getInboxTableItems(cmd,istore,olFolder,dasl,page,clear)
{
//get inbox items and add them to a store
//cmd=0 count, cmd=1 add items to store
//if page=0 add all records, otherwise go to specific page using jello.pageSize
//clear=true clear store before adding new records

//build query
lPage=0;
var iits;
if( dasl != null && dasl != "" && typeof(dasl) != "undefined")
	dasl = "@SQL="+dasl;
var colArray = emailColumns;
if (olFolder.Class==2)
{
var iits=getTable(olFolder,dasl);
if (olFolder.DefaultItemType != 0)
	colArray = taskColumns;
getFolderTable(iits,colArray);
iits.Sort(sortTables[olFolder.DefaultItemType],inboxSortOrder[olFolder.DefaultItemType]);
//if (dasl!=null && dasl != "")
//{var iits=ix.Restrict("@SQL="+dasl);}
//else
//{var iits=ix;}
}else{
var iits=olFolder;
// in code below remember that the item collection number from 1 not 0
if(olFolder.Columns.Item(2).Name == taskColumns[1])
	colArray = taskColumns;
//iits.Sort(colArray[1],true);
}
var counter=iits.GetRowCount();
if( counter == 0)
	return counter;



//getFolderTable(iits,colArray);




//add actions
if (cmd==1){addTableItemsToStore(iits,istore,page,clear,false, true);}

return counter;
}

function addTableItemsToStore(olitms,istore,page,clear,updateButtons, setlastfolder, isInboxStore)
{//add outlook item set in store
//alert(isInbox +" "+ page  +" "+ updateButtons);


if(isInboxStore == null || typeof(isInboxStore) == "undefined")
	isInboxStore = true;
status="Connecting and Calculating...";
var colArray = emailColumns;
try{var g=Ext.getCmp("igrid");g.hide();}catch(e){}
if (isInbox==true && page==1 && updateButtons==true && isInboxStore){pInbox();return;}
if (olitms==null)
{olitms=lastInboxItems;istore=lastListStore;}
//else if( setlastfolder == true && isInboxStore)
else if (typeof(Ext.getCmp("gotoStore")) == "undefined")
{lastInboxItems=olitms;lastListStore=istore;}

// in code below remember that the item collection number from 1 not 0
if(olitms.Columns.Item(2).Name == taskColumns[1])
	colArray = taskColumns;

//clear store
if (clear==true){istore.removeAll();}


	var psize=0;
	psize=parseInt(jello.pageSize);
	if( page <= 0) page = 1;
	// each row has an item values for one attribute of the table
	// for example, row ) might be EntryID of all items.
	var allCount = olitms.GetRowCount();
	if (page>0)
	{
		if( (page * psize) > (parseInt((allCount + psize -1) /psize)*psize))
			page = parseInt((allCount + psize -1) /psize);
		lStart=(psize*page)-psize;
		lEnd=(lStart+psize);}
	else{lStart=0;lEnd=(allCount> psize? psize:allCount);}
	try{Ext.getCmp("inboxgotopage").setRawValue(""+page);}catch(e){}
	if (lStart<0){lStart=0;}

	if (lEnd>allCount){lEnd=allCount;}
	var iArray = getTableArray(olitms,lStart, lEnd);
	if (allCount>0)
		addTableToStore(emailStoreObjNames,iArray, colArray, istore, 0, lEnd-lStart);
	else{}

	lPage=page;
if(isInboxStore){

if (isInbox==true ){allCount+=inboxTasksCount();}
if (updateButtons==true){updatePagingToolbar(istore,allCount);}
try{g.show();g.getView().focusRow(0);}catch(e){}
}
status=txtReady;
}

function addTableToStore(objNames,tabarray, colArray,xstore, start, end)
{
	//  cal with xstore null to return array and not add to store
	var recArray=[];
	var alerted = false;
	for (var x=start;x<end;x++)
	{
	var ar={};
	// use xw1 to index table of items as we can skip some cols in the colarray f they are blank
	var xw1=0;
	for( var xw = 0; xw < colArray.length; xw++)
	{
		if( colArray[xw] == "")
			continue;
		var cname = objNames[xw];
		if( cname == "iclass"){
			ar[cname] = mtypeToClass(tabarray[xw1++][x]);
		}else
			ar[cname] = tabarray[xw1++][x];
		if( typeof(ar[cname]) == "date") ar[cname] = new Date(String(ar[cname]));
	}
	// set attachment list e,pty - will be filled in if item is selected in view
	ar["attachmentList"] = "";
	// importance isn't item value is 1 of high else empty
	ar["importance"] = (ar["importance"] == 2? 1: "");
	// build the icon string from other elements
	ar["icon"] = ar["iclass"] + (ar["iclass"] == 43? ";0;0" : ";"+ ar["unread"] + ";" + ar["sensitivity"]);
	ar["groupon"] = (ar["iclass"] == 43? 1:0);
	var newRec=new actionRecord(ar);
	if( ar["subject"] == "" || ar["subject"] == null  || typeof(ar["subject"]) == "undefined")
		ar["subject"] = txtNoSubject;
	//if(! alerted && !ar[emailStoreObjNames[1]])
	//{alert("! match " + ar["subject"] + "  .. " + Ext.isDate(ar[emailStoreObjNames[1]])); alerted = true;}
	if( xstore != null){
		var sitems = xstore.query("entryID", ar.entryID);
		if (sitems.getCount() == 0)
			//xstore.addSorted(newRec);
			recArray.push(newRec);
	}else
		recArray.push(newRec);
	status =	txtStatusGetting+"..." + ar["subject"];
	}
	if( xstore == null)
		return recArray;
	xstore.add(recArray);
	var sstate = xstore.getSortState();
	xstore.sort(sstate.field,sstate.direction);


}

function updatePagingToolbar(listStore,counter,customCounterLabel)
{
//render the inbox items paging toolbar
try{
var ret="";var showNewer=false;var showOlder=true;
if (lPage==0){lPage++;}
var psize=parseInt(jello.pageSize);
if (lPage>1){showNewer=true;}

if (isInbox==false)
{if (psize*lPage>=counter){showOlder=false;}}
else
{if (psize*lPage>=(counter-inboxTasksCount())){showOlder=false;}}

if (showNewer==true){Ext.getCmp("inewer").enable();}else{Ext.getCmp("inewer").disable();}
if (showOlder==true){Ext.getCmp("iolder").enable();}else{Ext.getCmp("iolder").disable();}
if (isInbox==true){var ibc=inboxTasksCount();lStart=lStart+ibc;lEnd=lEnd+ibc;}
var ccls=lStart;if (lPage<2){ccls=lStart;}
ccls++; // add 1 since lstart is a zero based index
var ccle=lEnd;if (lEnd>counter){ccle=lEnd;}
if (customCounterLabel==null){customCounterLabel=txtDisplay+" " + ccls + "-" + (lEnd);}
//newfun
aposition.innerHTML="<font color=gray>"+lPage+" of "+parseInt(counter/psize)+"</font>";
updateCounter(customCounterLabel+" of "+counter);

}catch(e){}
}

function inboxTasksCount()
{
//count inbox tasks
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var dasl="urn:schemas-microsoft-com:office:office#Keywords IS NULL AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
var iits=iF.Restrict("@SQL="+dasl);
return iits.Count;
}

function filterInbox()
{//filter inbox items
	var lstore=Ext.getCmp("igrid").getStore();
	lPage=1;
var press=Ext.getCmp("ifilter").pressed;
	if (press==true)
	{
	var ft=prompt(txtPrmptFltVal,"");
	if (notEmpty(ft)==false){return;}
	Ext.getCmp("ifilter").setText(txtFilterOn+" ["+ ft +"]");
	var iits=lastInboxItems;
	var dasl="@SQL=urn:schemas:httpmail:subject LIKE '%"+ft+"%' OR urn:schemas:httpmail:fromname LIKE '%"+ft+"%' OR urn:schemas:httpmail:fromemail LIKE '%"+ft+"%' OR urn:schemas:contacts:cn LIKE '%"+ft+"%'";
	var its=iits.Restrict(dasl);
  its.Sort(sortTables[its.Parent.DefaultItemType],inboxSortOrder[its.Parent.DefaultItemType]);
//	lastInboxItems=its;
	isInbox=false;
	var counter=getInboxItems(1,lstore,its,null,1,true);
	updatePagingToolbar(lstore,counter);
  }
	else
	{Ext.getCmp("ifilter").setText("Filter");pInbox(lIbFolder);}
}

function getSecuredDataFromOutlook(olid)
{
//Trying to make user experience more complete, get sender and body information with one click and one security alert for the whole list
var extractOL = function(r){
      var itid=r.get("entryID");
      var it=NSpace.GetItemFromID(itid);
      var prw="";
			if (jello.mailPreviewHTML==true){prw=it.HTMLBody;}else{prw=cleanBody(it.HTMLBody);}

      r.set("body","<div>"+prw+"</div>");
            if (it.Class==43){r.set("sender",it.SenderName);r.set("to",it.To);}
            else
            {prw=it.Body;
            r.set("body","<div>"+cleanBody(prw)+"</div>");
            }

  };

   var ig=getActiveGrid();
   var ds=ig.getStore();

   if (olid==null)
   {ds.each(extractOL);}
   else
   {
   var ix=ds.find("entryID",olid);
   var r=ds.getAt(ix);
   extractOL(r);
   }

	try{
   var rs=new Array(ig.getSelectionModel().getSelected());
   ig.getSelectionModel().selectRecords(rs);
   }catch(e){}
}

function msgToolbar(isOLView)
{

var comboMax=300;
if (isOLView){comboMax=100;}
//render the inbox toolbar
    var dateMenu = new Ext.menu.DateMenu({
        handler : function(dp, idate){
            //var dt=idate.format('m/j/Y');
            //if(jello.dateFormat == 0 || jello.dateFormat == "0"){dt=idate.format('j/m/Y');}
            var dt=idate.toUTCString();
            inboxAction(null,'due',dt);
        }
    });

        var repmenu = new Ext.SplitButton({
        icon:'img\\reply.gif',
        cls:'x-btn-icon',
        tooltip: txtReplyInfo,
        id:'ireply',
		    handler:inboxAction,
        menu:new Ext.menu.Menu({
        items: [
        {
        icon: 'img\\list_review.gif',
        cls:'x-btn-icon',
        id:'repall',
        text: txtReplyAll,
		  	handler:inboxAction
      }]})
        });


    var delgmenu = new Ext.menu.Menu({
        id: 'dlgmenu',
        items: [{cls:'x-btn-icon',
        icon: 'img\\user.gif',
        text: txtDelgForInfo+' (Ctrl+E)',
    		handler:inboxAction
},{
        cls:'x-btn-icon',
        icon: 'img\\list_review.gif',
        text: txtRevForInfo,
        id:'irev',
		handler:inboxAction
			},{
        cls:'x-btn-icon',
        icon: 'img\\list_someday.gif',
        text: txtIncubForInfo,
			  id:'iinc',
		    handler:inboxAction
        }]});

       var actmenu = new Ext.menu.Menu({
        id: 'actmenu',

        items: [
		{
        icon:'img\\move.gif',
        cls:'x-btn-icon',
        text: txtMoveInfo + " (shift+ctrl+v)",
        id:'imove',
		handler:inboxAction
        },{
                cls:'x-btn-icon',
                text:txtCrAppForInfo,
                icon: 'img\\appoint.gif',
                handler:function(){
        inboxAction(null,'cal');}
        },        {
                text: txtTaskCnvToPrj,
                cls:'x-btn-icon',
                id:'aproject',
                icon: 'img\\project.gif',
                handler:function(){
        inboxAction(null,'prj');}
        }

        ]});

    tbar1 = new Ext.Toolbar();
    //tbar1.render('main');
        tbar1.add({
        icon: 'img\\new.gif',
        cls:'x-btn-icon',
        id: 'inew',
        tooltip: txtNewHere,
			listeners:{click: function(b,e){
			addOLItemHere(isInbox);}}
        },
        /*
        {
        iconCls: 'x-tbar-page-prev',
        tooltip: txtPrevious+' (j)',
        id:'ipit',
			listeners:{click: function(b,e){
			gridGo(-1);}}
        },{
        iconCls: 'x-tbar-page-next',
        tooltip: txtNext+' (k)',
        id:'init',
			listeners:{click: function(b,e){
			gridGo(1);}}
        }
        */
        '-',
        {
        menu:actmenu,
                listeners:{
        menushow: function(c){lowerOLViewCtl(true);},
        menuhide: function(c){lowerOLViewCtl(false);}
        },
        text:'Actions'
        },repmenu
        ,{
        icon:'img\\forward.gif',
        cls:'x-btn-icon',
        tooltip: txtFwdInfo,
		id:'iforw',
		handler:inboxAction
        },{
        icon: 'img\\page_delete.gif',
        cls:'x-btn-icon',
        tooltip: txtDelete+' (DEL)',
			listeners:{click: function(b,e){
			inboxAction(null,'del');}}
        }
       ,{
        tooltip: txtOpenInOutlook + '(Ctrl+Q)',
        cls:'x-btn-icon',
        icon: 'img\\info.gif',
        id:'iol',
		handler:inboxAction
        },'-',
        new Ext.form.ComboBox({
                fieldLabel: 'Tags',
                id: 'itags',
                store:globalTags,
                hideTrigger:false,
                typeAhead:false,
                width:100,
                listWidth:200,
                triggerAction:'all',
                emptyText:txtTagSelInfo,
                mode:'local',
                maxHeight:comboMax,
                listeners:{
         expand: function(c){lowerOLViewCtl(true);}
         ,
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                inboxAction(null,"tag");}
                setTimeout(function(){
            Ext.getCmp("itags").focus();
            },300);
                },

         collapse: function(c){lowerOLViewCtl(false);}
         ,
            select: function(cb,rec,idx){
            inboxAction(null,"tag");
            }
		}

            }),

		{
        iconCls:'olarchive',
        text:'<b>'+txtArchive+'</b>',
        tooltip: txtArcvInfo+' (Ctrl+A)',
        width:40,
        id:'iarc',
		handler:inboxAction
        },{
        cls:'x-btn-icon',
        icon: 'img\\collect16.png',
        id:'arconly',
        tooltip:txtArchiveOnly+' (Ctrl+B)',
			handler:inboxAction
      },'-',{
                cls:'x-btn-icon',
                icon: 'img\\refresh.gif',
                tooltip:txtRefresh+' (Ctrl+S)',
                id:'irefresh',
                handler:refreshInboxView
        },/*{
        icon: 'img\\filter.gif',
        cls:'x-btn-icon',
        tooltip: 'Set A View',
        id:'setvw',
        listeners:{click:function(){setActionView();}}
        }*/
		'-',{
    xtype: 'textfield',
		width:130,
		emptyText:'Add Action',
		hidden:isOLView,
		id:'inlinenewaction',
		listeners:{
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                quickAddAction(true);}},
        focus: function(t){t.getEl().fadeIn();}
		}},'-',{
                cls:'x-btn-icon',
                icon: 'img\\calendar.gif',
                tooltip:txtDueForInfo+' (Ctrl+D)',
                id:'popdate',
                menu: dateMenu
        },'-'
        ,new Ext.Toolbar.SplitButton({
        cls:'x-btn-icon',
        icon: 'img\\user.gif',
        tooltip: txtDelgForInfo+' (Ctrl+E)',
        id:'ideleg',
        menu:delgmenu,
        listeners:{
        menushow: function(dgbut,dgmen){lowerOLViewCtl(true);},
        menuhide: function(dgbut,dgmen){lowerOLViewCtl(false);}
        },
		handler:inboxAction
        })

        ,'-'
        );
return tbar1;

}

function msgBBar(grp)
{
//bottom toolbar

tbar2 = new Ext.Toolbar({
 //id:'gridfooter',
items:[
        {
        text:"Group List",
	    id:'igrouping',
	    tooltip:'Toggle grouping on/off',
	    enableToggle:true,
	    pressed:grp,
		listeners:{
            click: function(b,e){regroupInbox();}
            }
            },'-',{
        text:txtFilter,
	    id:'ifilter',
	    tooltip:txtFltrSenderIfo,
	    enableToggle:true,
		listeners:{
            click: function(b,e){filterInbox();}
            }
            }
			,{
        text:txtProps,
	    id:'fprops',
	    tooltip:txtFldrPropIfo,
	    enableToggle:false,
		listeners:{
            click: function(b,e){editOLFolder(fprops);}
            }
        },'-',{
        text:"<< "+txtNewer,
	    id:'inewer',
	    tooltip:'Ctrl+J',
	    disabled:true,
		listeners:{
            click: function(b,e){addToStore(null,null,(lPage-1),true,true);}
            }
            },'-',
new Ext.form.Label({
				html:'',
				id:'aposition'
				})
            ,'-',{
        text:txtOlder+" >>",
	    id:'iolder',
	    tooltip:'Ctrl+K',
	    disabled:true,
		listeners:{
            click: function(b,e){addToStore(null,null,(lPage+1),true,true);}
            }
             },'-',{
		text: "Go To Page : ",
		id: "gotolabel",
		listeners:{
            click: function(b,e){
			var val = Ext.getCmp('inboxgotopage').getRawValue();
      if( val <= 0 || val == "" || val == null || typeof(val) == "undefined")
					return;
			addToStore(null,null,val,true,true);}
            }


            },
            {
		xtype: 'textfield',
		width:50,
		//hidden:isOLView,
		id:'inboxgotopage',
		listeners:{
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
				var val = parseInt(f.getRawValue());
				if( val <= 0 || val == null || typeof(val) == "undefined")
					return;
                addToStore(null,null,val,true,true);}},
        focus: function(t){t.getEl().fadeIn();}
		}},
            {
        text:txtVwMailIbx,
        iconCls:'olinbox',
	    id:'ivolinbox',
	    tooltip:txtVwMailIbxIfo,
	    hidden:true,
		listeners:{
            click: function(b,e){pOLInbox(NSpace.GetFolderFromID(jello.inboxFolder));}
            }
            }

	] });


   return tbar2;
}


function gridGo(num)
{
//go to next/previous item on grid
var grd=getActiveGrid();
var g=grd.getSelectionModel();
if (g.hasSelection()==false){g.selectFirstRow();}
else{if (num>0){g.selectNext();}else{g.selectPrevious();}}
grd.getView().refresh();
}

function moveMailTo(olitem,tofolder,rec,noRec)
{
//move email to a folder and remove it from grid
var nit=olitem.Move(tofolder);
if (noRec){return;}
var store=Ext.getCmp("igrid").getStore();
gridGo((jello.goNextOnDelete?1:-1));
var grd=getActiveGrid();
var g=grd.getSelectionModel();
var r =g.getSelected();
var idx = store.indexOf(r);
store.remove(rec);
setBodyPopDisplay(idx);
Ext.getCmp("igrid").render();

}

function copyMailTo(olitem,tofolder,rec,noRec)
{
//move email to a folder and remove it from grid
var t = tofolder.DefaultItemType;
if( t == 0){
	var copyItem = olitem.Copy();
	copyItem.Move(tofolder);
	return;
}
if( t != 1 && t != 4 && t != 5 && t != 3)
	return;
var newA = tofolder.Items.Add();
try{
if (t != 5)
	newA.Subject = olitem.Subject;
newA.Body = olitem.Body;
newA.Categories = olitem.Categories;
newA.Display();
}catch(e){newA.Delete();}
}

function inboxAction(btn,vl,param,oneitem)
{
//execute inbox toolbar

var dpass = new Array();
refreshWholeView=false;
var itemsToArchive=new Array();
var citx;
if( typeof(oneitem) != "undefined" && oneitem != null)
	cit = oneitem;
else
	cit=getCheckedItems(null,param);
if (cit==false){return;}
var cl=cit.length;
var g=getActiveGrid();
try{if(param.substr(0,4)=="olv:"){g=null;}}catch(e){}
var noRecord=false;
if (typeof(g)=="undefined" || g==null){noRecord=true;}else{var store=g.getStore();}

var itemList=new Array();
var actionsCount=0;
var imsg="";var amsg="";

if (cl==0){return;}
var t=null;
	for (var x=0;x<cl;x++)
	{
		if (noRecord)
		{var it=cit[x];var id=it.EntryID;}
		else
		{
			var id=cit[x].get("entryID");
			try{
				var it=NSpace.GetItemFromID(id);
			}catch (e){store.remove(cit[x]);continue;}
		}

	if (btn!=null)
	{

 
	if (btn.id=="iarc" || btn.id=="arconly")
	//archive
	{
	itemsToArchive.push({olitem:it,record:cit[x],update:noRecord});
	}
	if ((btn.id=="ireply" || btn.id=="icmreply") && (it.Class==43 || it.Class==45))
	//reply
	//{var nit=it.Reply;nit.Display();}
		replyToMessage(it, true);
	if ((btn.id=="repall" || btn.id=="icmrepall") && (it.Class==43 || it.Class==45))
	//reply to all
	//{var nit=it.ReplyAll;nit.Display();}
		replyToMessage(it,false);
	if ((btn.id=="iforw" || btn.id=="icmforw") && (it.Class==43 || it.Class==45))
		//forward
		forwardMessage(it);
	if (btn.id=="irev")
	//review
	{toTagID(it,cit[x],7,noRecord);imsg=txtMsgTagReview;}
	if (btn.id=="ideleg")
	//delegate
	{
	itemList.push(cit[x]);
	}
	if (btn.id=="iinc")
	//incubate
	{toTagID(it,cit[x],9,noRecord);imsg=txtMsgIncubate;}
	if (btn.id=="iol" || btn.id=="icmol")
	//open outlook item
	{it.Display();}
	if( btn.id == "icont" || btn.id=="icmcont")
		showSenderContact(id);
	if ((btn.id=="imove" || btn.id=="icmmove") && it.Class==43)
	{//move

		if (jello.useGoToFolderToMove==false || jello.useGoToFolderToMove==0 || jello.useGoToFolderToMove=="0")
		{
			 if (t==null){t=NSpace.PickFolder();}
				 if (t!=null)
				 {
					 if (t.DefaultItemType==0){

						 try{moveMailTo(it,t,cit[x],noRecord);}catch(e){alert(txtMsgNoMoveErr);return;}
						 if (!noRecord){try{updateRecordItem(cit[x],true);}catch(e){}}
					 }else
						 try{copyMailTo(it,t);}catch(e){alert(txtMsgNoMoveErr);return;}

					 imsg=txtMsgMoveTo+" <a class=jellolink style=text-decoration:underline; onclick=outlookView('"+t.EntryID+"')><b>"+t+"</b></a>";

				 }

		}
		else
		{
	try{var xlist = [it.EntryID, cit[x], noRecord];
	dpass.push(xlist);} catch(e){}
		}

	}
	}
		if (vl=="tag")
		{//tag item
    var ctlname="itags";
		  if (notEmpty(param)){var p=param.split(":");ctlname+=":"+p[1];}
    var tg=Ext.getCmp(ctlname).getValue();
	    if (notEmpty(tg)==false)
		{var tg=Ext.getCmp(ctlname).getRawValue();}
    if (notEmpty(tg)){if (addTagTo(tg,it,cit[x],noRecord)==false){return;}}
		if (it.Class==48){actionsCount++;}
		}

    		if (vl=="prj")
		{//convert action to project
			if (it.Class==48)
			{
			lastOpenTagID=0;
			var aFailed=convertToProject(id);
			if (!noRecord){updateRecordItem(cit[x],true);}
			}else{alert(txtInvalid);return;}
		}


		if (vl=="del" || vl=="icmdel")
		{//delete item
		if (t==null){t=confirm(txtMsgDelItem);}
		if (t==true){OLDeleteItem(id);
			if (!noRecord){updateRecordItem(cit[x],true);gridGo((jello.goNextOnDelete?1:-1));}
			imsg=txtMsgDeleted;}
		else{imsg=null;}
		}

		if (vl=="due" || vl=="icmdue")
		{//add due date

					if( !setDueDate(it,param))
				if (!noRecord){updateInboxRecord(cit[x]);}

		}

		if (vl=="cal")
		{//create calendar entry w/ item(s) as text
			//the appointment will be created in user's calendar folder respecting settings
    	if (jello.createAppointmentsOnDefCal==0 || jello.createAppointmentsOnDefCal=="0")
      {var cf=NSpace.GetFolderFromID(jello.calendarFolder);}
      else
      {var cf=calendarItems;}

		t=cf.Items.Add();
	  t.Subject=it.Subject;
	  try{t.Body=it.Body;}catch(e){}
	  t.Display();
		}

	}

// do deletes and moves -
try{
  	if( dpass.length > 0){selFolder(txtAToFolder, moveOrCopyMessage, dpass,false);}
	}catch(e){}

var msg="";
	try{if (btn.id=="ideleg"){contactList=new Array();delegateItems(itemList);}}catch(e){}
		try{
		if (btn.id=="iarc"){msg=doArchive(itemsToArchive);}
		if (btn.id=="arconly"){msg=doArchive(itemsToArchive,null,true);}
		}catch(e){}

//user info messages
if (vl=="due" || vl=="icmdue"){param="";}
var arcLink="<a class=jellolinktop onclick=archiveSelecions('"+param+"');><b>"+txtArchive+"</b></a>";

		if (vl=="tag")
		{
		    if (actionsCount==cl){msg=actionsCount+" "+txtMsgArcOK.replace("%1",arcLink);}
		if (cl>actionsCount){msg=cl+" "+txtMsgArcTskOK.replace("%1",arcLink);}
		    if (typeof(tg)!="undefined"){
		    Ext.info.msg(txtTag+' '+tg,msg);
		    var tg=Ext.getCmp(ctlname).reset();
		    }
		}

		if (vl=="due" || vl=="icmdue")
		{
		if (actionsCount==cl){msg=actionsCount+" "+txtMsgArcNDueOK.replace("%1",arcLink);}
		if (cl>actionsCount){msg=cl+" "+txtMsgArcTskDueOK.replace("%1",arcLink);}
		Ext.info.msg(txtDueDateOK,msg);
		}

		if ((vl=="del" || vl=="icmdel") && imsg!=null)
		{Ext.info.msg(txtMsgActionComp,cl+" "+txtItemItems+" "+imsg);}


	try
	{
		if (btn.id=="irev" || btn.id=="iinc" /*|| btn.id=="imove"*/)
		{Ext.info.msg(txtMsgActionComp,cl+" "+txtItemItems+" "+imsg);}
		if (btn.id=="iarc" || btn.id=="arconly")
		{Ext.info.msg(txtMsgTtlArcv,msg);}
	}catch(e){}

//if there is an outlook view control update toolbar counter
if (refreshWholeView){pInbox();return;}
updateOLViewCounter();
//if did archive clear selections
try{
if (btn.id=="iarc" || btn.id=="arconly"){
  if (jello.selectFirstItem==0 || jello.selectFirstItem=="0")
  {
  g.getSelectionModel().clearSelections();
  }
}
}catch(e){}
setTimeout(function(){updateTheLatestThing();},6);
status=txtReady;
}

function OLDmoveOrCopyMessage(folderID, data, makeTask)
{
	// maketask will be true if the checkbox was checked on the move dialogue

	var t = NSpace.GetFolderFromID(folderID);
	var tasks = new Array();
	for( n = 0; n < data.length; n++)
	{
		var item = data[n];
		var  it = NSpace.GetItemFromID(item[0]);
		var cit = item[1];
		var noRecord = item[2];
		if (t.DefaultItemType==0){
			if( typeof(makeTask) != "undefined" && makeTask != null && makeTask == true){
				//copyMailTo(it,NSpace.GetFolderFromID(jello.actionFolder));
				var ta = new Array();
				ta.push({folder:t});
				try{if(it.Categories == "") it.Categories = "!Next";}catch(e){}
				tasks.push(archiveItem(it,cit,ta,noRecord,true,true)[0]);
			}
			else
				try{moveMailTo(it,t,cit,noRecord);}catch(e){alert(txtMsgNoMoveErr);continue;}
			if (!noRecord){try{updateRecordItem(cit,true);}catch(e){}}
		}else
			try{copyMailTo(it,t);}catch(e){alert(txtMsgNoMoveErr);continue;}

	}
	var imsg=txtMsgMoveTo+" <a class=jellolink style=text-decoration:underline; onclick=outlookView('"+t.EntryID+"')><b>"+t+"</b></a>";
	Ext.info.msg(txtMsgActionComp,data.length +" "+txtItemItems+" "+imsg);
	if( makeTask && tasks.length == 1 && tasks[0].status == true){
		var xarray = new Array();
	    var xao = new actionObject(tasks[0].task);
	    var xar = new actionRecord(xao);
	    xarray.push(xar);
	    editAction(xarray, true,false);
	}


}

function moveOrCopyMessage(folderID, data, actions)
{
	//actions obj has elements task, appt, reply.  True if action should be performed

	var t = null;
	if( folderID != null && folderID != "" )
		t = NSpace.GetFolderFromID(folderID);
	var tasks = new Array();
	for( n = 0; n < data.length; n++)
	{
		var item = data[n];
		var  it = NSpace.GetItemFromID(item[0]);
		var cit = item[1];
		var noRecord = item[2];
		if (t == null || t.DefaultItemType==0){
			if( actions.reply == true) replyToMessage(it, true);
			if( actions.fwd == true){forwardMessage(it);}
			if( actions.appt == true) {
				var aptt=calendarItems.Items.Add();
				aptt.Subject=it.Subject;
				try{aptt.Body=it.Body;}catch(e){}
				aptt.Display();
			}
			if( actions.task == true){
				//copyMailTo(it,NSpace.GetFolderFromID(jello.actionFolder));
				var ta = new Array();
				if( t != null) ta.push({folder:t});
				try{if(it.Categories != "")
					it.Categories += ";";
				it.Categories +=  actions.cats;
				if( it.Categories == "")
					it.Categories = "!Next";
				}catch(e){}
				tasks.push(archiveItem(it,cit,ta,(t == null? false:noRecord),false,true)[0]);

			}
			if( actions.copy == true){
				try{if(t != null)copyMailTo(it,t);}catch(e){}
			}
			if( actions.del  == true){
				if( actions.task != true){
					it.Delete();
					if(!noRecord){try{updateRecordItem(cit,true);}catch(e){}}
				}
			}
			else if( actions.copy != true)
				try{if(t != null)moveMailTo(it,t,cit,noRecord);}catch(e){alert(txtMsgNoMoveErr);continue;}
			if (t != null && actions.copy != true && !noRecord){try{updateRecordItem(cit,true);}catch(e){}}
		}else{
			try{if(t != null)copyMailTo(it,t);}catch(e){alert(txtMsgNoMoveErr);continue;}
			if( actions.del  == true){
				it.Delete();
				if(!noRecord){try{updateRecordItem(cit,true);}catch(e){}}

			}
		}

	}
	var imsg = "";
	if( t != null) imsg = txtMsgMoveTo+" <a clatss=jellolink style=text-decoration:underline; onclick=outlookView('"+t.EntryID+"')><b>"+t+"</b></a>";
	Ext.info.msg(txtMsgActionComp,data.length +" "+txtItemItems+" "+imsg);
	if( actions.task == true && tasks.length == 1 && tasks[0].status == true){
		var xarray = new Array();
	    var xao = new actionObject(tasks[0].task);
	    var xar = new actionRecord(xao);
	    xarray.push(xar);
	    editAction(xarray, true,false);
	}
	//var nInItems = inboxTable.GetRowCount();
	//inboxICount.innerHTML = "("+ nInItems.toString() +")";

}


function archiveSelecions(param)
{//helper archive function to run from hyperlinks
inboxAction(Ext.getCmp('iarc'),null,param);
}




function addTagTo(tagname,olitem,rec,noRec)
{//add a tag to an inbox item
if (olitem.Class==44){Ext.info.msg('Alert', txtInvalid);return;}
tagname=Trim(tagname);
var r=getTag(tagname);
var wasNew=false;

if (typeof(r)=="undefined")
{
//if tag not exists create it
var choice=false;
if (choice=confirm(txtMsgNewTag.replace("%1",tagname))==false){return false;}
if (notEmpty(tagname)==false){return;}
createTag(tagname);
wasNew=true;
}

addOLItemCategory(olitem,tagname);
if (noRec){return;}
updateInboxRecord(rec);
	if (wasNew)
	{
	setTimeout(function(){refreshView();},6);
	}

}
function renderInboxSubject(v,m,r)
{
//render the subject line in grid
var tg=r.get("tags");
var id=r.get("entryID");
var st=r.get("status");
var dd=r.get("due");
var cl=r.get("iclass");
var un=r.get("unread");
var dw=r.get("markfordownload");
var rt="";
var tList=tagList(null,"[INBOX]",tg,id);
if (isDate(dd))
{
var idate=new Date(dd);
if (idate.getYear()<4500)
{
var disdd=DisplayDate(idate);
tList[0]="<a class=jellolink title='"+txtDtSetTo+":"+disdd+" ["+txtClickEdit+"]' onclick=dueDateProperties('"+id+"','"+disdd+"')><span style='background-image:url("+imgPath+"calendar.gif);background-repeat:no-repeat;padding-left:18px;width:18;height:16;'></span></a>"+tList[0];}
}
  var subj=Ext.util.Format.htmlEncode(v);
  if (un==true){subj="<b>"+subj+"</b>";}
  if (tList[1]==true){subj="<b>"+subj+"</b>";}
  if (dw==4){subj="<strike>"+subj+"</strike>";}
  rt+=tList[0]+"&nbsp;"+subj;
  rt=NextActionToggleIcons(tg,id)+rt;
 return rt;
}

function updateInboxRecord(rec)
{//update grid record after change
try{var id=rec.get("entryID");}catch(e){return;}
var g=getActiveGrid();
var store=g.getStore();
var idx=store.indexOf(rec);
var g=g.getSelectionModel();
var wasSel=g.isSelected(rec);
store.remove(rec);
var it=NSpace.GetItemFromID(id);
var ar=new actionObject(it);
var newRec=new actionRecord(ar);
store.insert(idx,newRec);
if (wasSel==true){var nrr=new Array();nrr.push(newRec);g.selectRecords(nrr,true);}
}

function removeInboxTag(tg,id)
{
//remove tag by x from inbox view
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
var g=getActiveGrid();
  if (g!=null)
  {
    var store=g.getStore();
    var ix=store.find("entryID",id);
    var rec=store.getAt(ix);
    updateInboxRecord(rec);
  }  
}

function openInboxItem(id)
{//open inbox item for editing
if( typeof(id) == "undefined" || id == null || id == ""){
var g=Ext.getCmp("igrid").getSelectionModel();
var r=g.getSelected();
var xid=r.get("entryID");
}else
	var xid = id;
try{var it=NSpace.GetItemFromID(xid);}catch(e){alert(txtOItemDeleted);return;}
if (it.Class==48){var t=new Array(r);g.selectRecords(t);editAction(t);}
else{try{it.Display();}catch(e){alert(txtOItemDeleted);}}
}

// function getJDueProperty(it){
// //get custom due field value for emails. Create if not exists
// var jn="";
// try{
// jn=it.UserProperties.Item("jdue").Value;
// }catch(e){it.UserProperties.Add("jdue",1,true);}
// return jn;
// }

function getJDueProperty(it){
//get custom due field value for emails. Create if not exists
var jn="";
try{
	jn=it.UserProperties.Item("jdue").Value;
	}catch(e){ if( OLversion <12 || jello.useOLDue == false) it.UserProperties.Add("jdue",1,true);}
if(OLversion >= 12 && jello.useOLDue){
// for ol 12 use the taskduedate field
// convert if custom field already in use
	var dn =  it.TaskDueDate;
	if( jn != ""  && isDate(jn) ){
		it.itemProperties.item("TaskDueDate").Value = DisplayDate(new Date(jn));
		try{it.UserProperties.Item("jdue").Value = "";}catch(e){}
	}else{
		jn = dn;
	}
}
return jn;
}

function setJDueProperty(it,due)
{
	if( OLversion < 12 || jello.useOLDue == false)
	{
		if( due == "1/1/4501")
			due = "";
		try{
			it.UserProperties.Item("jdue").Value = due;
		}catch(e){ it.UserProperties.Add("jdue",1,true);
			it.UserProperties.Item("jdue").Value = due;
		}
	}else{
		it.itemProperties.item("TaskDueDate").Value = due;
	}
}

function archiveItem(it,rec,copyToo,noRec,archiveIt,createTask)
{
//archiveItem(it,rec,af,noRec,archiveIt,createTask,copyToo);
//it=outlook item
//rec=grid/store record
//copyToo=Array of archive folders. Move to first, copy to others
//noRec=no Parent store record to update t/f
//archiveIt=Original item to be archived? t/f
//createTask=Create task linked with the arc item? t/f

//return array:1.Status success/fail, 2.Created Task ref, 3.Archived Item ref

	if (!noRec)
	{
	var g=getActiveGrid();
	if(typeof(g)=="undefined"){noRec=true;}
	}

	 var cats="";var ibod="";var idue;var newA=null;var noUpdate=false;var astatus=true;

	//get item's categories
    if (it.Class!=44)
    {
    cats=it.itemProperties.item(catProperty).Value;
    if (notEmpty(cats)){cats=checkForNotExistingTags(cats);}else{createTask=false;}
    }

	    //get due date of item
	if (noRec)
	{
	try{idue=getJDueProperty(it);}catch(e){idue=null;}
	}else
	{
	try{idue=rec.get("due");}catch(e){}
	}

	//if item has a due date a task has to be created despite tagging
	if (isDate(idue) && notEmpty(idue))
		{var idd=new Date(idue);if (idd.getYear()<4000 && it.Parent.EntryID!=jello.actionFolder){createTask=true;}}

    //get body of item
    if (!noRec){try{ibod=rec.get("body");}catch(e){}}
    else{ibod=it.Body;}

	if (createTask)
	{
		if (ibod.substr(0,9)=="<secured>")
		{try{ibod=it.Body;}catch(e){}
	}else{
		var tl=toplayer.innerHTML;toplayer.innerHTML=ibod;ibod=toplayer.innerText;toplayer.innerHTML=tl;}
	}


	var finalItem;
	if (archiveIt)
	{
	//move the item to the first archive folder
		if (copyToo.length>0)
		{
		var af=copyToo[0].folder;
		if (it.Parent==af){finalItem=it;}else{
		try{finalItem=it.Move(af);}catch(e){alert(txtArchiveCannot);return;}
		}
		}
	}
    else
    {finalItem=it;noUpdate=true;if (it.Class==48){noUpdate=false;}}

		if ((notEmpty(cats) || isDate(idue)) && createTask )
		{
		//create task connected to this item if needed
		newA=createActionOL(finalItem.Subject);newA.Categories=cats;

			if (notEmpty(idue)){
				var idd = new Date(idue);
				if (idd.getYear()<4000 )
          if(typeof(idue)=="string")
          {newA.itemProperties.item(dueProperty).Value=idd.format('Y/m/d');}
          else
          {newA.itemProperties.item(dueProperty).Value=idue.format('Y/m/d');}

			}

			//attach original item to created task
      setAttachmentProperty(newA,finalItem.EntryID);
      newA.Save();
      setAttachmentProperty(finalItem,newA.EntryID);
      finalItem.Save();

	//update new task's body to original item;s body
	  if (notEmpty(ibod))
	  {
	  newA.Body=ibod;
	  if (jello.autoUpdateTaskNotes==false)
	  {getJNotesProperty(newA);setJNotesProperty(newA,ibod);}
	  }
	  //check if original item must be copied to other locations as well
		if (copyToo !=null && copyToo.length>1)
		 {
	  for (var x=1;x<copyToo.length;x++)
	  {
		var copyToFolder=copyToo[x].folder;
				  if (copyToFolder==finalItem.Parent)
				  {var itemCopy=finalItem;}
					 else
				  {
				  var icp=finalItem.Copy();
				  var itemCopy=icp.Move(copyToFolder);
				  setAttachmentProperty(newA,itemCopy.EntryID,true);
				  newA.Save();
		  setAttachmentProperty(itemCopy,newA.EntryID);
				  itemCopy.Save();
				  }

		newA.Save();
            }
         }
    	newA.Save();
		}
		else
		{
    //no task creating case or other types
      if (it.Class==48 && !notEmpty(cats)){noUpdate=true;astatus=false;}
    }


		//update store if there is one
		if(!noRec && !noUpdate)
		{
		
		var store=g.getStore();
		try{gridGo((jello.goNextOnDelete?1:-1));}catch(e){}
		store.remove(rec);
		try{updateAFooterCounter(store);}catch(e){}
		}

		if (noUpdate && !noRec)
		{
    var store=g.getStore();
		updateRecordItem(rec,false);
    }

    updateTheLatestThing();
var res=new Array();
res.push({status:astatus,task:newA,archivedItem:finalItem,noArc:noUpdate});
return res;
}

function getDefaultArcFolder()
{//return the default archive folder if not exists prompt for it
	if (notEmpty(jello.archiveFolder))
	{
	var arf=setAndCheckArcFolder(jello.archiveFolder);
	if (arf!=null){return arf.EntryID;}
	}

//if no archive folder defined define one
var choice=false;
if (choice=confirm(txtPrmptNoArcFldr)==false){return;}
var t=NSpace.PickFolder();
	if (t!=null )
	{
	if (t.EntryID==jello.inboxFolder){alert(txtMsgNoSystemFolder);return;}
		if (t.DefaultItemType==0)
		{
		//must be a mail folder
			var stt="jello.archiveFolder='"+t.EntryID+"'";
			eval(stt);
			jese.saveCurrent();
			var arf=setAndCheckArcFolder(jello.archiveFolder);
			return arf.EntryID;
		}
		else
		{
		alert(txtPromptMailFolder);
		return null;
		}
	}
}
function toTagID(it,rec,tid,noRec)
{
//add review tag to item & archive
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
var rtt=r.get("tag");
addOLItemCategory(it,rtt);
var ita=new Array();
ita.push({olitem:it,record:rec,update:noRec});
var newtask=doArchive(ita,true);
return newtask;
}

function countInboxItems()
{//inbox counter
var counter=0;
try{
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var dasl=getInboxTaskDASL();
//var dasl="urn:schemas-microsoft-com:office:office#Keywords IS NULL AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
var iiF=iF.Restrict("@SQL="+dasl);
counter=iiF.Count;
	if (jello.noUseMail==false || jello.noUseMail==0 || jello.noUseMail=="0")
	{
iF=NSpace.GetFolderFromID(jello.inboxFolder).Items;
counter+=iF.Count;
	}
}catch(e){}
return counter;
}

function addOLItemHere(isInbox,inFolder)
{
//Add a new item of the default type
var nit;

  if (isInbox!=null)
  {
  if (isInbox || (lIbFolder==null)){var nit=NSpace.GetFolderFromID(jello.inboxFolder).Items.Add();}
  else{if(folderOnView==0){folderOnView=NSpace.GetFolderFromID(jello.inboxFolder);}nit=folderOnView.Items.Add();}
  }
  else
  {
  var f=NSpace.GetFolderFromID(inFolder);
  nit=f.Items.Add();
  }

nit.Display();
}


function checkArchiveFolder(it)
{
//check if an archive folder has been defined for tags present or globally. If none ask for one

//first check tags of item (if its an email)
var afolder=new Array;
  if (it.Class>0)
  {
  var cats=it.itemProperties.item(catProperty).Value;
    cats=cats.replace(new RegExp(",","g"),";");
    var c=cats.split(";");
    if (c.length>0)
    {
          for (var x=0;x<c.length;x++)
          {
          var af=getTagArcFolder(Trim(c[x]));
            if (notEmpty(af))
            {
				if(!valExistsInArray(af,afolder)){afolder.push(af);}
            }
            else
            {
            //if no custom folder for tag add default arc folder
            var daf=getDefaultArcFolder();
				if (notEmpty(daf))
				{
				if(!valExistsInArray(daf,afolder)){afolder.push(daf);}
				}
            }
          }
    }

  if (afolder.length>0){return afolder;}
  }

return null;
}

var inboxUpdateInterval = 60000;

var inboxTask = {
    run: function(){
        getNewInboxItems(false);
    },
    interval: inboxUpdateInterval //1 minute
};
var inboxUpdateRunner;

var tabReader = new Ext.data.ArrayReader({}, [
		{name: 'subject'},
        {name: 'entryID'},
		{name: 'importance'},
		{name: 'sender'},
		{name: 'to'},
		{name: 'cc'},
		{name: 'body'},
		{name: 'attachment'},
		{name: 'created',type:'date'},
		{name: 'due',type:'date'},
		{name: 'unread',type:'boolean'},
		{name: 'attachmentList'},
		{name: 'iclass'},
		{name: 'icon'},
		{name: 'tags'},
		{name: 'groupon'},
		{idProperty: 'entryID'}
]);

function pInbox(folder)
{
lastContext="";
// if was inbox, stop the auto-updater
if((isInbox ||(folder != null && folder.entryID == jello.inboxFolder))&& typeof(inboxUpdateRunner) != "undefined" && typeof(inboxUpdateRunner) != "null")
	inboxUpdateRunner.stop(inboxTask);

initScreen(true,"pInbox("+(typeof(folder)=="undefined"?"":folder)+")");
thisGrid="igrid";
//Inbox management process
lIbFolder=folder;

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
		{name: 'type'},
		{name: 'tags'},
		{name: 'groupon'}
]);

if (folder != null){if (folder.DefaultItemType>0){useTables=false;}}

var listStore = new Ext.data.GroupingStore({
        id:'actStore',
		reader: ((OLversion >=12 && useTables)?tabReader:ibreader),
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
		   {name: 'tags'},
		   {name: 'groupon'}

        ],
		groupField: "groupon",
		groupMode: 'display',
		showGroupName: false,
		sortInfo:{field: 'created', direction: "DESC"}
    });

getTagsArray();
//main.innerHTML="";

var tasksCounter=0;
var counter=0;
if (folder==null)// || (folder.entryID == jello.inboxFolder))
{//inbox view
isInbox=true;
if (OLversion >=12 && useTables)var iBoxArray = getTableArray(inboxTable,0,25);
var iF=NSpace.GetFolderFromID(jello.actionFolder);
var dasl=getInboxTaskDASL();
if (jello.noUseMail==1 || jello.noUseMail=="1"){showNoMailInbox(dasl);return;}

counter=getInboxItems(1,listStore,iF,dasl,-999,true); //get all tasks

tasksCounter=counter;
if (jello.noUseMail==false || jello.noUseMail==0 || jello.noUseMail=="0")
  {
  folder=setAndCheckArcFolder(jello.inboxFolder);
  if (folder==null){alert(txtMsgInvInbFol);}
  else{
  	  lastInboxUpdateTime = new Date();
	//counter+=getInboxItems(1,listStore,folder,null,1,false);} // get first page of emails
	counter+=getInboxItems(1,listStore,folder,folder.CurrentView.Filter,1,false);
  }
  
  }
}
else
{//any folder view
isInbox=false;
folderOnView=folder;

//does user want an native Outlook view?
var nativeView=false;
var ix=olFolderStore.find("eid",new RegExp("^"+folder.EntryID+"$"));
var r=olFolderStore.getAt(ix);
  if (r!=null)
  {nativeView=r.get("activex");}
   //Outlook 2010 can have native views only in IE or hta mode
   if (OLversion>13 && conStatus=="Outlook"){nativeView=false;}
if (nativeView)
{
    pOLInbox(folder);return;
}

//counter=getInboxItems(1,listStore,folder,null,1,true);
var tFilter="";
 
if (OLversion>11)
      {
      if (folder.Store.IsDataFileStore)
      {
      tFilter=folder.CurrentView.Filter;
      counter=getInboxItems(1,listStore,folder,tFilter,1,true);
      }
      else
      {
      counter=getInboxItems(1,listStore,folder,"",1,true);
      } 
      }else
      {
      tFilter=folder.CurrentView.Filter;
      counter=getInboxItems(1,listStore,folder,tFilter,1,true);
      }

}

var gridwaitMask = new Ext.LoadMask(Ext.getBody(), {msg:"Please wait..."},{store:listStore});
sm= new Ext.grid.CheckboxSelectionModel({});


var cc=listStore.getCount();

fprops="";
if (!isInbox){fprops=folderOnView.EntryID;}else{fprops=jello.inboxFolder;}


//if( !isInbox)
	//try{listStore.sort("created", (inboxSortOrder[folder.DefaultItemType]?"DESC":"ASC"));}catch(e){}

 var itbar=msgToolbar();
 var ibbar=msgBBar(jello.inboxGroupOn);
 var inboxChosenView=new Ext.grid.GroupingView({forceFit:true,
    // custom grouping text template to display the number of items per group
    groupTextTpl: '{[values.text.search("task.gif") > 0 ? "'+txtCbActions+': " : [values.text.search("mail.gif") > 0 ? "'+txtSetMessages+': " : "'+txtType+'"+([values.text]) ] ]}  ({[values.rs.length]} {[values.rs.length > 1 ? "Items" : "Item"]})'
    });

if (jello.inboxGroupOn==false){inboxChosenView=null;}

//Ext.getCmp("previewpanel").expand(true);
//Ext.getCmp("previewpanel").setHeight(jello.actionPreviewHeight);

    var grid = new Ext.grid.GridPanel({
        store: listStore,
        tbar:itbar,
        bbar:ibbar,
        id:'igrid',
        columns: [
        sm,
			{header: "", width: 20, fixed:true, sortable: true, renderer: getImportance, dataIndex: 'importance'},
			{header: "", width: 0, hidden: true, fixed:true, sortable: false, renderer: groupName, dataIndex: 'groupon'},
			{header: "", width: 35, groupable: ((OLversion >=12 && useTables)?false:true), groupName: "Type",hideable:false, fixed:true, sortable: true, renderer: getIcon, dataIndex: 'icon'},
            {header: txtDate, width: 80, hidden:false, sortable: true, renderer:DisplayInboxDate, dataIndex: 'created'},
            {header: txtSubject, width: 400, sortable: true, renderer: renderInboxSubject,dataIndex: 'subject'},
            {header: txtSender, width: 150, sortable: true, dataIndex: 'sender'}

        ],

     		view: inboxChosenView,
        stripeRows: true,
        autoScroll:true,
        deferRowRender:false,
        //split: true,
		region: 'north',
        enableColumnHide:true,
        viewConfig:{
        emptyText:txtNoDispItms
        },
        trackMouseOver:true,
        height:Ext.getCmp("centerpanel").getHeight()-jello.actionPreviewHeight-30,
		sm:sm,
    listeners:{
    mouseover: function(e){thisGrid='igrid';},
		rowdblclick: function(g,row,e){
			var rec =g.getStore().getAt(row);
			if( jello.doubleClickBodyPopup && rec.get("iclass") == 43){
				popTaskBody(rec.get("entryID"),rec);
			}else
				openInboxItem(null);
		},
       cellcontextmenu: function(g, row,cell,e){
       g.getSelectionModel().selectRow(row);
			rightClickItemMenu(e,row, g);
			}
		},
        enableColumnMove:true,
        border:false,
        loadMask:gridwaitMask,
//        autoHeight:true,
        layout:'fit'
        //bbar:tbar2

    });


    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.doLayout();


    grid.getSelectionModel().on('rowselect', function(sm, rowIdx, r) {
		var detailPanel = Ext.getCmp('previewpanel');
		detailPanel.rowidx = rowIdx;
		updateMailItemDisplay(r);
		mailTpl.overwrite(detailPanel.body, r.data);
		Ext.getCmp("previewpanel").setHeight(jello.actionPreviewHeight);
		var pop= Ext.getCmp("bodytxt");
		if( pop != null && typeof(pop) != "undefined")
			setBodyPopDisplay(rowIdx);
		/*
		//Ext.getCmp("previewpanel").doLayout();
				var gb=r.get("body");
		if (gb.substr(0,9)!="<secured>" && jello.markReadfromJello==true)
				{//if user can read preview update read property of message
				var itid=r.get("entryID");
					try{
					var it=NSpace.GetItemFromID(itid);
          it.UnRead=false;it.Save();
					}catch(e){}
				}*/
	   });

grid.on('columnresize',function(index,size){saveGridState("igrid");});
grid.on('sortchange',function(){saveGridState("igrid");});
grid.getColumnModel().on('columnmoved',function(){saveGridState("igrid");});
grid.getColumnModel().on('hiddenchange',function(){saveGridState("igrid");});
restoreGridState("igrid");
var pp= Ext.getCmp("previewpanel");

//update toolbar buttons for task only inbox
/*
if (jello.onlyTaskInbox==true || jello.noUseMail==true)
{try{
Ext.getCmp("iolder").hide();
Ext.getCmp("inewer").hide();
Ext.getCmp("aposition").hide();
Ext.getCmp("ifilter").hide();
Ext.getCmp("ivolinbox").show();
}catch(e){}
}  */


setTimeout(function(){
try{
if (jello.selectFirstItem==1 || jello.selectFirstItem=="1"){grid.getSelectionModel().selectRow(0);}
grid.getView().focusRow(0);grid.focus();}catch(e){}
resizeGrids();
},6);

//-----shortcut keys--------------
var gmap1 = new Ext.KeyMap('igrid', {
    key: 13,
    fn: function(){
    openInboxItem(null);
    },
    scope: this
});

//Ctrl+M moves item
var gmap2 = new Ext.KeyMap('igrid', {
    key: 'v',
    fn: function(){inboxAction(Ext.getCmp("imove"));},
    stopEvent:true,
    ctrl:true,
    shift:true,
    scope: this
});


//Ctrl+J/K moves to previous/next item
var gmap2 = new Ext.KeyMap('igrid', {
    key: 'k',
    fn: function(){gridGo(1);},
    stopEvent:true,
    ctrl:true,
    shift:false,
    scope: this
});

var gmap3 = new Ext.KeyMap('igrid', {
    key: 'j',
    fn: function(){gridGo(-1);},
    stopEvent:true,
    ctrl:true,
    shift:false,
    scope: this
});

//ctrl+alt+J/K moves to previous/next page
var gmap4 = new Ext.KeyMap('igrid', {
    key: 'k',
    fn: function(){if (Ext.getCmp("iolder").disabled==false){addToStore(null,null,(lPage+1),true,true);}},
    stopEvent:true,
    ctrl:true,
    alt:true,
    scope: this
});

var gmap5 = new Ext.KeyMap('igrid', {
    key: 'j',
    fn: function(){if (Ext.getCmp("inewer").disabled==false){addToStore(null,null,(lPage-1),true,true);}},
    stopEvent:true,
    ctrl:true,
    alt:true,
    scope: this
});

var gmap6 = new Ext.KeyMap('igrid', {
    key: 'a',
    fn: function(){inboxAction(Ext.getCmp('iarc'));},
    stopEvent:true,
    ctrl:true,
    scope: this
});

var gmap61 = new Ext.KeyMap('igrid', {
    key: 'b',
    fn: function(){inboxAction(Ext.getCmp('arconly'));},
    stopEvent:true,
    ctrl:true,
    scope: this
});



var gmap7 = new Ext.KeyMap('igrid', {
    key: 't',
    ctrl:true,
    fn: function(){Ext.getCmp('itags').focus();},
    stopEvent:true,
    scope: this
});

var gmap8 = new Ext.KeyMap('igrid', {
    key: 'd',
    ctrl:true,
    fn: function(){
	Ext.getCmp('popdate').showMenu();
            },
    scope: this});

var gmap9 = new Ext.KeyMap('igrid', {
    key: 's',
    ctrl:true,
    fn: function(){
	refreshInboxView();
            },
    scope: this});

var gmap10 = new Ext.KeyMap('igrid', {
    key: 'e',
    ctrl:true,
    fn: function(){
	inboxAction(Ext.getCmp('ideleg'));
   },
    stopEvent:true,
    scope: this});

//Ctrl+q open OL item
var gmap12 = new Ext.KeyMap('igrid', {
    key: 'q',
    fn: function(){
    inboxAction(Ext.getCmp('iol'));
    },
    stopEvent:true,
    ctrl:true,
    scope: this
});
 // ctl+i update since last update/orig display
var gmap14 = new Ext.KeyMap('igrid', {
    key: 'i',
    ctrl:true,
    fn: function(){
	getNewInboxItems(true);
   },
    stopEvent:true,
    scope: this
});

 // ctl+i update since last update/orig display
var gmap15 = new Ext.KeyMap('igrid', {
    key: 'w',
    ctrl:true,
    alt:true,
    fn: function(){
	ticklerPopup();
   },
    stopEvent:true,
    scope: this
});

//DEL deletes items
var gmap11 = new Ext.KeyMap('igrid', {
    key: Ext.EventObject.DELETE,
    fn: function(){
    inboxAction(null,'del');
    },
    stopEvent:true,
    ctrl:false,
    scope: this
});

//insert focuses the new action toolbar box
var gmap12 = new Ext.KeyMap('igrid', {
    key: Ext.EventObject.INSERT,
    fn: function(){Ext.getCmp("inlinenewaction").focus();try{var aw=Ext.getCmp("subj");aw.focus();}catch(e){}},
    stopEvent:true,
    ctrl:false,
    shift:false,
    scope: this
});


updatePagingToolbar(listStore,counter);//,txtDisplay+" " + (lStart+1) + "-" + (tasksCounter+lEnd));

if (isInbox==true)
{
try{
ppnl.setTitle("<img src=img\\inbox16.png style=float:left;> "+txtInbox+ "(" + folder.CurrentView.Name+")");
}catch(e){}
if (jello.inboxAutoRefresh==true || jello.inboxAutoRefresh==1 || jello.inboxAutoRefresh=="1"){
var inboxUpdateRunner = new Ext.util.TaskRunner();
inboxUpdateRunner.start(inboxTask);
}

}
else
{
try{
ppnl.setTitle("<img src=img\\"+icFolder+" style=float:left;> "+folder.FolderPath+ "(" + folder.CurrentView.Name+")");
}catch(e){}
}
//resizeGrids(Ext.getCmp("igrid"),100);
status=txtReady;
}

function groupName(c,m,r){
    if( c == 1)
		return txtSetMessages;
	else
		return txtTask;

}

function updateMailItemDisplay(r)
{
try{var it = NSpace.GetItemFromID(r.get("entryID"));
		var ats = it.Attachments;
		alist="";
		if( ats.Count > 0 && r.get("attachmentList") == ""){
			attachmentList = "";
			var colon = "";
			for( var i=1; i <= ats.Count; i++){
				var z = ats.Item(i);
				alist += colon + "&nbsp;<a onclick=olItem('"+it.EntryID+"'); class=jelloLink style='color:blue;text-decoration:underline;'>"+z.FileName+"</a>";
				colon=";";
			}
		}
		r.beginEdit();
		if( alist != "") r.set("attachmentList",alist);
		r.set("body",getItemBody(it));
		if( it.Subject != "")
			r.set("subject", it.Subject);
		r.endEdit();
		var gb=r.get("body");
		if (gb.substr(0,9)!="<secured>" && jello.markReadfromJello==true)
		{//if user can read preview update read property of message
          it.UnRead=false;it.Save();
		}
}catch(e){}
}
//var inboxUpdateInterval = 60000;
var iboxTimerCount=0;


function getNewInboxItems(keypress)
{
if (jello.noUseMail==1 || jello.noUseMail=="1"){return;}
if( OLversion >= 12 && useTables){getNewInboxTableItems(keypress);return;}
	status="Updating...";
	iboxTimerCount++;
	try{if(isInbox == false  || getActiveGrid().getId() != "igrid"){

		return;}
	}catch(e){return;}

		var nInItems = NSpace.GetFolderFromID(jello.inboxFolder).Items.Count;
	inboxICount.innerHTML = "("+ nInItems.toString() +")";


	try{
		if( lPage != 1)
			return;

		var tDate = lastInboxUpdateTime;
		if (tDate == null || tDate == "")
		{
			lastInboxUpdateTime = new Date();
			//status=txtReady;
      return;
		}
		var g = getActiveGrid();
		// get the store for items
		var istore=g.getStore();


		// to handle possible race conditions
		// set the minute back 1 minute
		tDate.setMinutes(tDate.getMinutes()-1);

		//var DRecv = (tDate.getMonth() +1) + "/" + tDate.getDate() + "/" + tDate.getFullYear() + " " + theHour + ":" + m.toString()  + ampm;
		var DRecv=tDate.getUTCFullYear() + "/" + (tDate.getUTCMonth()+1)+"/"+tDate.getUTCDate()+" "+tDate.getUTCHours()+":"+tDate.getUTCMinutes();
		var dasl = "@SQL=(urn:schemas:httpmail:datereceived >= '" + DRecv + "')";
		var folder=setAndCheckArcFolder(jello.inboxFolder);

		var iF=folder.Items;
		var iits=iF.Restrict(dasl);
		var iTime = new Date(); // remember time we got array
		// any items that are new?

		if( iits.Count == 0){


			//status=txtReady;
      return;
		}

		var nic=iits.Count;
		for (var x=1;x<=nic;x++)
		{
			try{ var ar=new actionObject(iits.Item(x));}catch(e){nic--;continue;}
			// check to see if we have a dup, is it in the store
			// needed because restrict granularity is minute, rounded down

			var sitems = istore.query("entryID", ar.entryID);
			if (sitems.getCount() != 0)
				continue;
			status=txtStatusGetting+" ("+x+"/"+nic+")..." + ar.subject ;
			var newRec=new actionRecord(ar);
			istore.addSorted(newRec);

		}

		lastInboxUpdateTime = iTime;
		//istore.sort("created", "DESC");

		//lastInboxItems=olitms;lastListStore=istore;
		var sm= g.getSelectionModel();
			sm.clearSelections();
			g.getView().refresh();
			g.getView().selectFirstRow();




//		status = "";


	}catch(e){}
//status=txtReady;
}
var lastInboxItemCount=0;
function getNewInboxTableItems(keypress)
{


	// for tables we can always have the items count
	var nInItems = inboxTable.GetRowCount();
	if( lastInboxItemCount != nInItems)
		{inboxICount.innerHTML = "("+ nInItems.toString() +")";
		lastInboxItemCount = nInItems;
		}
	// now check to see if we need to process
	iboxTimerCount++;
	var g = getActiveGrid();
	try{if(isInbox == false  || g.getId() != "igrid" || lPage != 1){

		return;}
	}catch(e){return;}

	try{
		status="Updating...";
		var tDate = lastInboxUpdateTime;
		if (tDate == null || tDate == "")
		{
			lastInboxUpdateTime = new Date();
			//status=txtReady;
			return;
		}

		// get the store for items
		var istore=g.getStore();


		// to handle possible race conditions
		// set the minute back 1 minute
		tDate.setMinutes(tDate.getMinutes()-1);

		var filterD = tDate.format(jelloDateFormatString(true) +" g:i A");
		var dasl = "[ReceivedTime] >= '" + filterD  + "'";
		// get time now to use as next base if query matches
		var iTime = new Date();
		// use filter to find new items, if any
		var newItemsTable = inboxTable.Restrict(dasl);
		var theCount ;
		if( (theCount = newItemsTable.GetRowCount()) == 0){
			status="";
			return;
		}
		// get the items and add to the store
		getFolderTable( newItemsTable,emailColumns);
		//newItemsTable.Sort("ReceivedTime");
		//addToStore(newItemsTable,istore,0, false,true, false);
		var iArray = getTableArray(newItemsTable,0, theCount);
		addTableToStore(emailStoreObjNames,iArray, emailColumns, istore, 0, theCount);
		// set new time to use as base for next updates
		lastInboxUpdateTime = iTime;
		g.render();
		status = "";
	}catch(e){}

}

function mailContextMenu(rec,id, it, inlist)
{

try{Ext.getCmp("inrclickmenu").destroy();}catch(e){}
try{Ext.getCmp("flagmenu").destroy();}catch(e){}


	return new Ext.menu.Menu({
		id: 'inrclickmenu',
		items: [

		{
        icon:'img\\reply.gif',
        text: txtReplyInfo,
        id:'icmreply',
		listeners:{click: function(b,e){
			mailClickForwarder(b,e,inlist,rec);}}
		},
		{
        icon: 'img\\list_review.gif',
        id:'icmrepall',
        text: txtReplyAll,
			listeners:{click: function(b,e){mailClickForwarder(b,e,inlist,rec);} }
		},

		{
        icon:'img\\forward.gif',
        text: txtFwdInfo,
		id:'icmforw',
		listeners:{click: function(b,e){mailClickForwarder(b,e,inlist,rec);} }
		},'-',{
	        icon:'img\\move.gif',
          text: txtMoveInfo,
        id:'icmmove',
		listeners:{click: function(b,e){mailClickForwarder(b,e,inlist,rec);}}
    },{
        icon: 'img\\page_delete.gif',
        text: txtDelete,
        id: 'icmdel',
		listeners:{click: function(b,e){
			mailClickForwarder(b,e,inlist,rec);}}
		},'-',
		{
		text: txtBody,
		id:'icmbody',
		handler:function(){
		popTaskBody(id,rec);}
		},
		{
            icon: 'img\\calendar.gif',
            text:txtDueForInfo,
            id:'icmpopdate',
            menu: new Ext.menu.DateMenu({
        	handler : function(dp, idate){
            //var dt=idate.format('m/j/Y');
            //if(jello.dateFormat == 0 || jello.dateFormat == "0"){dt=idate.format('j/m/Y');}
            var myb = new Ext.Button({id:"icmdue"});
            var dt=idate.toUTCString();
            mailClickForwarder(myb,null,inlist,rec,dt);
        	}
    		})
        },
		    {
        text: txtOpenInOutlook,
        icon: 'img\\info.gif',
        id:'icmol',
		listeners:{click: function(b,e){mailClickForwarder(b,e,inlist,rec);}}
		},{
        icon: 'img\\contact.gif',
        id: 'icmcont',
        text: txtConView,
		listeners:{click: function(b,e){mailClickForwarder(b,e,inlist,rec);} }
		},{
			text: txtShowWithTag,
			id: 'avsametasks',
			icon: 'img\\task.gif',
			listeners: {click: function(b,e){tasksWithTag(rec);} }
		}

		]
	});
}

function mailClickForwarder(b,e,inlist, rec,param)
{
	if( inlist){
		if( b.id == "icmdel" || b.id == "icmdue" )
			inboxAction(null,b.id,param);
		else
			inboxAction(b,e);
	}else{
		var x = new Array();
		x.push(rec);
		if( b.id == "icmdel" || b.id == "icmdue")
			inboxAction(null,b.id, param ,x);
		else
			inboxAction(b,null,null,x);
	}
}


function updateOLViewCounter()
{
 try{updateCounter(olnative.ItemCount);}catch(e){}
 //it would be good to update the pagin toolbar after archiving
 //try{var str=Ext.getCmp("igrid").getStore();updatePagingToolbar(str,str.getCount(),null);}catch(e){}
}

function getInboxTaskDASL()
{
//catch tasks: 1.Category is null, 2. Not Completed, 3. Not having a jello tag
var ds=tagStore;
ds.filter("istag",true);

var dasl="";
var addToSQL = function(r){
var ctx=r.get("tag");
if (notEmpty(dasl)){dasl+=" AND ";}
dasl+="(NOT urn:schemas-microsoft-com:office:office#Keywords LIKE '%" + ctx + "%')";
  };

ds.each(addToSQL);
dasl="(("+dasl+")";
dasl+=" OR (urn:schemas-microsoft-com:office:office#Keywords IS NULL)) AND (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2)";
ds.clearFilter();
return dasl;
}

function doArchive(aItems,returnTask,arcOnly)
{
//manage archive for a list of items
var ic=aItems.length;var cTasks=0;
var aCounter=0;var archivers=new Array();var newTasks=new Array();var noCounter=0;
if (ic==0){return;}

	for (var x=0;x<ic;x++)
	{
	var i=aItems[x];
	var it=i.olitem;var rec=i.record;var noRec=i.update;var otype=it.Class;var createTask=false;var archiveIt=false;
	//support for meeting request, responses etc.
  if (otype==53 || otype==54 || otype==181 || otype==55 || otype==56 || otype==57){otype=1000;}
	   switch (otype)
     {
     case 43:archiveIt=true;createTask=true;break;
     case 45:archiveIt=true;createTask=true;break;
     case 48:archiveIt=false;createTask=false;cTasks++;break;
     case 1000:archiveIt=true;createTask=true;break;
     default:archiveIt=false;createTask=true;break;
     }

	//if arcOnly is true do not create task (archive only button)
    if (arcOnly==true){createTask=false;}

	var arcfolder=checkArchiveFolder(it);
	var copyToo=new Array();
	if (arcfolder!=null)
	{
		if (arcfolder.length==1)
		{//one archive folder. Do archive
		  var af=setAndCheckArcFolder(arcfolder[0]);
		  copyToo.push({folder:af});
		  var ait=archiveItem(it,rec,copyToo,noRec,archiveIt,createTask);
		  if(!valExistsInArray(af.EntryID,archivers)){archivers.push(af.EntryID);}

        }
		else
		{
    //found multiple folders. copy item to all
		//all copies will be linked to the created task
  			for (var x=0;x<arcfolder.length;x++)
  			{
  	      var cf=setAndCheckArcFolder(arcfolder[x]);
          copyToo.push({folder:cf});
          if(!valExistsInArray(cf.EntryID,archivers)){archivers.push(cf.EntryID);}
    		}
      var ait=archiveItem(it,rec,copyToo,noRec,archiveIt,createTask);
		}

		//process the result
		//{status:true,task:newA,archivedItem:finalItem}
      if (ait[0].status==true)
		  {
			aCounter++;
      //archivers.push(af);
			if (ait[0].task!=null){newTasks.push(ait[0].task);}
	        }
		  else
		  {
		  noCounter++;
		  }
		}
		else
		{
		var ait=archiveItem(it,rec,null,noRec,false,false);
		if (ait[0].status==true){aCounter++;}else{noCounter++;}
		}
	}

      if (returnTask){return ait[0].task;}

    var msg="";	var arcAll="";
    if (aCounter==1)
    {
    var aai=ait[0].archivedItem;
    arcAll=archiversList(archivers);
    msg="<b>"+aai+"</b> "+txtArcToIfo+" "+arcAll;
    if (ait[0].noArc){msg=txtProcessed;}
    if (aai.Parent.EntryID==jello.actionFolder){msg=txtTkFiledIfo;}
    if (ait[0].task!=null){msg+="<br>"+txtAndCrThisSub+" <a class=jellolinktop onclick=scAction('"+ait[0].task.EntryID+"')>"+txtAction+"</a>";}
    return msg;
    }
    if (aCounter>0)
    {
    arcAll=archiversList(archivers);var ams=txtProcessed;
    if (arcAll.length>1 && cTasks==0){ams=txtArcToIfo+" "+arcAll;}
    if (cTasks>0 && (aCounter-cTasks)>0){ams=cTasks+" "+txtProcessed+" "+(aCounter-cTasks)+" "+txtArcToIfo+" "+arcAll;}
    msg=aCounter+" "+txtItemItems+" "+ams;
    if (noCounter>0)
    {
    msg+=" and <b>"+noCounter+"</b> "+txtItemItems+" "+txtNoProcessIfo;}

    var xarray = new Array();
    var xao = new actionObject(ait[0].task);
    var xar = new actionRecord(xao);
    xarray.push(xar);
    editAction(xarray, true,false);

    return msg;
    }
    if (noCounter>0){msg+=txtMsgNoProcess;return msg;}

}


function archiversList(archivers)
{
//create a list of folder in which items were archived for the msgbox
var ret="";
		if (archivers.length>0)
		{
			for (var x=0;x<archivers.length;x++)
			{	try{
			var ff=NSpace.GetFolderFromID(archivers[x]);
				ret+="<a class=jellolinktop onclick=OLFolderOpenNewWindow('"+archivers[x]+"')>"+ ff +"</a> ";
				}catch(e){}
			}
		}
return ret;
}


function setAndCheckArcFolder(id)
{//return a valid outlook archive folder ot null
var rt;
try{rt=NSpace.GetFolderFromID(id);}catch(e){rt=null;}
return rt;
}


function DisplayInboxDate(t,m,r)
{
//display of date and time for calendar views
var sd=new Date(r.get("created"));
if (isNaN(sd)){return "";}
var dt2 = new Date().add(Date.DAY, -6);
var dayOnly = false;
if (sd.between(dt2, new Date()))
	dayOnly = true;
var tmm=sd.format("H:i");
var ret=DisplayAppDate(sd,dayOnly);
ret+=" <span class=caltime>("+tmm+")</span>";
return ret;
}

function savePrePanelHeight(pname)
{
var hght = (Ext.getCmp(pname).getInnerHeight())+25;
if( pname == 'igrid')
jello.mailPreviewHeight = hght;
if( pname == 'grid')
jello.actionPreviewHeight = hght;
jese.saveCurrent();
}

function prePanelAddResizeEvent(gname)
{
 var pp=Ext.getCmp("prepanel");
 pp.on('resize',function(){savePrePanelHeight(gname);});
}

function refreshInboxView()
{
var lastFolder=lIbFolder;
pInbox(lastFolder);
}

function dueDateProperties(id,cdate)
{
try{
var it=NSpace.GetItemFromID(id);

}catch(e){alert(txtInvalid);return;}

          var msg=it+"<br>"+txtDueDate+":<b>"+cdate+"</b>";
          var buts={no:txtDuePropsRemove,cancel:txtCancel};

          Ext.Msg.show({
           title:txtDuePropsEdit,
           msg: msg,
           buttons: buts,
           fn: function(b,t){duePrpFn(b,it);},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });

}

function duePrpFn(b,it)
{
//handle due date of items based on choice

    if (b=="no")
    {
    //remove due

		if (it.Class!=48)
			{

			var dp=getJDueProperty(it);
			setJDueProperty(it,"1/1/4501");it.Save();
			}
		else
			{
			it.itemProperties.item(dueProperty).Value="1/1/4501";it.Save();
			}
	var id=it.EntryID;
    var s=Ext.getCmp("igrid").getStore();
    var ix=s.find("entryID",id);
    var r=s.getAt(ix);
	if (r!=null){updateRecordItem(r,false);}
    }


}
function showSenderContact(id)
{
	if(Ext.isEmpty(id)){
		Ext.MessageBox.alert(txtContact,txtContactNo);
		return;
	}
	try{
	var it=NSpace.GetItemFromID(id);
	var xregex = new RegExp("XXX","g");
	var sender = it.SenderEmailAddress;
	var iF = NSpace.GetFolderFromID(jello.contactFolder);
	var basedasl1 = "urn:schemas:contacts:email1 = 'XXX' OR urn:schemas:contacts:email2 = 'XXX' OR urn:schemas:contacts:email3 = 'XXX'";
	var basedasl2 = " urn:schemas:contacts:cn =  'XXX'";
	var dasl = basedasl1.replace(xregex,sender) + " OR " + basedasl1.replace(xregex,sender.toLowerCase()) + " OR " + basedasl2.replace(xregex,it.SenderName);

	var match = iF.Items.Restrict("@SQL="+dasl);
	// now try for lastname
	if( match.Count == 0) {
		// no space than just return
		var ix;
		var ssender = it.SenderName;
		var sname="";
		var sname2="";
		if( (ix= ssender.indexOf('@') ) != -1)
			var sname = ssender.substring(0,ix);
		if( (ix= ssender.indexOf(' ') ) != -1)
			sname2 = ssender.substring(ix+1);
		if( sname == "" && sname2 == ""){
			Ext.MessageBox.alert(txtContact,txtContactNo);
			return;
		}
		var dasl = "";
		if( sname != "")
			dasl = basedasl1.replace(xregex,"%"+sname+"%") + " OR " + basedasl2.replace(xregex,"%"+sname.toLowerCase()+"%");
		var dasl2="";
		if( sname2 != "")
			dasl2 = basedasl2.replace(xregex,"%"+sname2+"%") + " OR " + basedasl2.replace(xregex,"%"+ sname2.toLowerCase()+"%");
		if( dasl != ""){
			if( dasl2 != "")
				dasl += " OR " + dasl2;
		}else
			dasl = dasl2;
		var xreg2 = new RegExp("=","g");
		dasl = dasl.replace(xreg2, "LIKE");
		//dasl = "([Email1Address] like '%" + sname + "%' OR [Email2Address] like '%" + sname + "%' OR [Email3Address] like '%" + sname + "%' ";
		//var sname1 = sname.toLowerCase();
		//dasl += " OR [Email1Address] like '%" + sname1 + "%' OR [Email2Address] like '%" + sname1 + "%' OR [Email3Address] like '%" + sname1 + "%' ";
		//dasl += "  \"urn:schemas:contacts:cn\" LIKE '%" + sname + "%' OR \"urn:schemas:contacts:cn\" LIKE '%" + sname + "%')";
		match = iF.Items.Restrict("@SQL="+dasl);
		if( match.Count == 0){
			Ext.MessageBox.alert(txtContact,txtContactNo);
			return;
		}
	}
	if( match.Count == 1){
		var xit = match.Item(1);
		xit.Display();
	}else{
	var xcont = new Array();
	for(var i=1; i <= match.Count ; i++ )
	{
		var xit = match.Item(i);
		xcont.push(new Array(xit.FullName + "("+ xit.Email1Address + ")", xit.EntryID));
	}
	selContact(xcont);

	}
 }catch(e){Ext.MessageBox.alert(txtContact,txtContactNo);}
}
function selContact(cdata)
{
	var cnameStore = new Ext.data.SimpleStore({
				fields: ['namedisp','id'],
				data : cdata
			});
	var combo = new Ext.form.ComboBox({
			id: "selFoldCombo",
			store: cnameStore,
			anyMatch: true,
			matchWordStart: false,
			queryIgnoreCase: true,
			displayField:'namedisp',
			selectOnFocus: true,
			typeAhead: true,
			mode: 'local',
			//emptyText:txtGotoInfo+'...',
			selectOnFocus:true,
			listeners : {
				select : function(combo,record,idx){
					var id = record.get("id");
					//open outlook view
					var myit = NSpace.GetItemFromID(id);
					myit.Display();
					win.destroy();
			},
			forceSelection: true
		}});
	var win = new Ext.Window({
        title: txtContact,
        width: 300,
        height: 100,
        id:'selectcont',
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
  setTimeout(function(){Ext.getCmp("selFoldCombo").focus(true,100);},1);

	//shortcut keys
var map = new Ext.KeyMap('selectcont', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});
}

function showNoMailInbox(dasl)
{
//display the non-mail inbox for the case user wants to manage only untagged actions in Inbox
showContext(-1,1,dasl);
lastContext="##Inbox";
}

// create a reply, if check is false, always reply to all if multiple folks
function replyToMessage(it, check)
 {
		var xans = check;
		var nit;
		var ic = it.To.split(";");
		if( xans && (it.CC != "" || ic.length > 1))
			xans = confirm(txtConfirmReply);
		if( xans)
			nit = it.Reply();
		else
			nit = it.ReplyAll();

		//var nit=it.Reply;
		
		nit.Display();
 }
 function forwardMessage(it)
 {
	var nit = it.Forward();
	nit.Display();
 }
 
 function regroupInbox(a,b)
 {//enable/disable grouping of items
 
 var press=Ext.getCmp("igrouping").pressed;
 jello.inboxGroupOn=press;
 jese.saveCurrent();
 pInbox();
 }