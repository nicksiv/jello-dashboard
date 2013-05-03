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


//define ticklerRecord
 var ticklerRecord = Ext.data.Record.create([
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
	{name: 'done'},
	{name: 'duration'},
	{name: 'tags'}

]);

var reader = new Ext.data.ArrayReader({}, [
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
	{name: 'done'},
	{name: 'duration'},
	{name: 'sortDate'},
	{name: 'icon'},
	{name: 'tags'},
	{name: 'tag'}
]);

var thisRangeStart;
var thisRangeEnd;
var thisCalView;
var totalDuration;

function pTicklers(periodToStore,popup)
{
if( typeof(popup) == "undefined" ||  popup != true)
initScreen(false,"pTicklers("+periodToStore+","+popup+")");
thisGrid="tgrid";

var listStore = new Ext.data.GroupingStore({
reader: reader,
sortInfo:{field: 'due', direction: "ASC"},
id:'tickStore',
groupField:'daypos',
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

 var counter=0;

	if (notEmpty(periodToStore))
	{
	//return only the resulting store
		switch(periodToStore)
		{
		case "D":
			thisRangeStart=new Date();thisRangeEnd=new Date();break;
		case "W":
			var now =  new  Date();
			var nowDayOfWeek = now.getDay()-1;var nowDay = now.getDate();
			var nowMonth = now.getMonth();var nowYear = now.getYear();
			nowYear += (nowYear < 2000) ? 1900 : 0;
			thisRangeStart = new Date(nowYear, nowMonth, nowDay - nowDayOfWeek);
			thisRangeEnd = new Date(nowYear, nowMonth, nowDay + (6 - nowDayOfWeek));
			break;
		case "M":
			thisRangeStart.setDate(1);var ldom=daysInMonth(thisRangeStart.getMonth()+1,thisRangeStart.getFullYear());
			thisRangeEnd.setDate(ldom);
			break;
		case "A":
			thisRangeStart=new Date();thisRangeEnd=new Date();
			thisRangeEnd.setTime(thisRangeStart.getTime() + (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));
			break;
		}

	getTicklerFile(thisRangeStart,thisRangeEnd,1,listStore,true);
	return listStore;
	}

//create grid
var gridwaitMask = new Ext.LoadMask(Ext.getBody(), {msg:txtMsgWait+"..."},{store:listStore});
sm= new Ext.grid.CheckboxSelectionModel({});

/*
 tbar2 = new Ext.Toolbar({
 id:'gridfooter',
items:[
new Ext.form.Label({
				html:counter + " " + txtItemItems,
				id:'acounter'
				})
	] });
  */
var dts=jello.dateSeperator;
var fmt="n"+dts+"j"+dts+"y";
var ttbar=ticklerToolbar();

    var grid = new Ext.grid.EditorGridPanel({
        store: listStore,
        id:'tgrid',
		clicksToEdit: 1,
        tbar:ttbar,
        columns: [
        sm,
			{header: "", width: 15, fixed:true, sortable: true, renderer: getImportance, dataIndex: 'importance'},
			{header: "", width: 30, hideable:false, fixed:true, sortable: false, renderer: getIcon, dataIndex: 'icon'},
            {header: txtTickler, width: 400, sortable: true, renderer: renderTickSubject,dataIndex: 'subject'},
            {header: txtStartDate, width: 170, hidden:false, sortable: true, renderer:DisplayCalDate, dataIndex: 'due',
				editable: true,editor: new Ext.form.DateField({
                    format: fmt,
                    editable: false
                })
			},
            {header: txtDuration, width: 70, hidden:false, sortable: true, renderer:DisplayDuration, dataIndex: 'duration'},
            {header: txtLocation, width: 170, hidden:false, sortable: true, dataIndex: 'location'},
            {header: txtPosition, width: 0, hidden:true, sortable: true, dataIndex: 'daypos'}
        ],
        stripeRows: true,
        autoScroll:true,
        deferRowRender:false,
        enableColumnHide:true,
           view: new Ext.grid.GroupingView({
            forceFit:true,
            groupTextTpl: '{text} ({[values.rs.length]} {[values.rs.length > 1 ? "Items" : "Item"]})'
        }),

        viewConfig:{
        emptyText:txtNoDispItms
        },
        trackMouseOver:true,
        height:Ext.getCmp("centerpanel").getHeight()-jello.actionPreviewHeight-30,
		sm:sm,
    listeners:{
		mouseover: function(e){thisGrid='tgrid';},
		rowdblclick: function(g,row,e){
		openTicklerItem();
        },
        cellcontextmenu: function(g, row,cell,e){
			rightClickItemMenu(e,row,g);
		},
		afteredit: function (obj){
					actionDateEdited(obj);
		},
		beforeedit : function(obj)
		{
			var typ = obj.record.get("iclass");
			if (typ !=48 &&  typ != 43)
				return false;
			else
				return true;
		 }

		},
        enableColumnMove:true,
        border:false,
        //width:panelWidth-5,
        loadMask:gridwaitMask,
        layout:'fit'
        //bbar:tbar2

    });

    //main.innerHTML="<div id=toolbar></div><div id=list></div>";
    //grid.render('list');
    //ticklerToolbar();
    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.setAutoScroll(false);
    ppnl.doLayout();

    grid.on('columnresize',function(index,size){saveGridState("tgrid");});
    grid.on('sortchange',function(){saveGridState("tgrid");});
    grid.getColumnModel().on('columnmoved',function(){saveGridState("tgrid");});
    grid.getColumnModel().on('hiddenchange',function(index,size){saveGridState("tgrid");});
    restoreGridState("tgrid");


		switch(jello.defaultCalendarView)
		{
		case "0":ticklerRangeSelected(Ext.getCmp("vday"));thisCalView="vday";break;
		case "1":ticklerRangeSelected(Ext.getCmp("vweek"));thisCalView="vweek";break;
		case "2":ticklerRangeSelected(Ext.getCmp("vmonth"));thisCalView="vmonth";break;
		case "3":ticklerRangeSelected(Ext.getCmp("vagenda"));thisCalView="vagenda";break;
		}

if (notEmpty(thisCalView)){Ext.getCmp(thisCalView).toggle(true);}
updateTicklerTitle();

setTimeout(function(){
try{
if (jello.selectFirstItem==1 || jello.selectFirstItem=="1"){grid.getSelectionModel().selectRow(0);}
grid.getView().focusRow(0);grid.focus();}catch(e){}
resizeGrids();
},6);

//ctrl + INS adds appointment
var tmap1 = new Ext.KeyMap('tgrid', {
    key: Ext.EventObject.INSERT,
    fn: function(){ticklerSelected(Ext.getCmp("vnew"),null);},
    ctrl:false,
    shift:false,
    stopEvent:true,
    scope: this
});

//DEL deletes items
var tmap2 = new Ext.KeyMap('tgrid', {
    key: Ext.EventObject.DELETE,
    fn: function(){
    ticklerSelected(Ext.getCmp('vdel'));
    },
    stopEvent:true,
    ctrl:false,
    scope: this
});

//ENTER opens items
var tmap3 = new Ext.KeyMap('tgrid', {
    key: Ext.EventObject.ENTER,
    fn: function(){
    openTicklerItem();
    },
    stopEvent:true,
    ctrl:false,
    scope: this
});

//ctrl+q open in outlook
var tmap3 = new Ext.KeyMap('tgrid', {
    key: 'q',
    ctrl:true,
    fn: function(){
	ticklerInOutlook();
    },
    stopEvent:true,
    scope: this
});
}

function getTicklerFile(fDate,tDate,fun,store,clear,isMove)
{
//get outlook dated items in range and update a store
//fun 0=counter only
totalDuration=0;
try{var g=Ext.getCmp("tgrid");g.hide();}catch(e){}
fDate.setHours(00,00,00);
tDate.setHours(23,59,59);
var originalFDate=fDate;
//if user does not want to include past due appointments start from now
  if ((jello.hidePastAppointments==true || jello.hidePastAppointments==1) && isMove!=true)
  {fDate=new Date();}
var counter=0;
if (clear==true){store.removeAll();}
//set the calendar items query
var DSstart=fDate.getUTCFullYear() + "/" + (fDate.getUTCMonth()+1)+"/"+fDate.getUTCDate()+" "+fDate.getUTCHours()+":"+fDate.getUTCMinutes();
var DSend=tDate.getUTCFullYear() + "/" + (tDate.getUTCMonth()+1)+"/"+tDate.getUTCDate()+" "+tDate.getUTCHours()+":"+tDate.getUTCMinutes();
var dasl="@SQL=(urn:schemas:calendar:dtstart <= '"+DSend+"' AND urn:schemas:calendar:dtend  >= '"+DSstart+"')";
var iF=NSpace.GetFolderFromID(jello.calendarFolder).Items;
iF.Sort("urn:schemas:calendar:dtstart");
iF.IncludeRecurrences = 1;
var its=iF.Restrict(dasl);


var Ap = its.GetFirst();
//loop through occurences
	do
	{

		if (Ap!=null)
		{
      	if (fun>0){
					if (Ap.BillingInformation!="[done]")
					{
					var ce=new ticklerObject(Ap);
					var newRec=new ticklerRecord(ce);
					store.add(newRec);
					addDuration(newRec);
					status=txtStatusGetting+"..." + Ap;
					counter++;
					}
					else
					{
						if (jello.CalendarShowCompleted=="1" || jello.CalendarShowCompleted==true)
						{
						var ce=new ticklerObject(Ap);
						var newRec=new ticklerRecord(ce);
						store.add(newRec);
						addDuration(newRec);
						status=txtStatusGetting+"..." + Ap;
						counter++;
						}
					}
				  }

		}

	Ap = its.GetNext();
	}while(Ap!=null);
//status=txtReady;
//get Tasks with due dates
var DSstart=originalFDate.getUTCFullYear() + "/" + (originalFDate.getUTCMonth()+1)+"/"+originalFDate.getUTCDate()+" "+originalFDate.getUTCHours()+":"+originalFDate.getUTCMinutes();
var doneFilter="AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003<>2";
if (jello.CalendarShowCompleted=="1" || jello.CalendarShowCompleted==true){doneFilter="";}

var dasl="@SQL=";
if (jello.allODueTasks=="1" || jello.allODueTasks==1)
{
//show all overdue tasks  and tasks in period
dasl+="http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2 AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040<'Today'";
dasl+="OR (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040 >='"+DSstart+"' and http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040 <='"+DSend+"') "+doneFilter;
dasl+=" OR (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040>='"+DSstart+"' and http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040<='"+DSend+"') "+doneFilter;

}
else
{
//show tasks in period
dasl+="(http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040 >='"+DSstart+"' and http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040 <='"+DSend+"') "+doneFilter;
dasl+=" OR (http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040>='"+DSstart+"' and http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040<='"+DSend+"') "+doneFilter;
}


var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
iF.Sort("http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81040040");

var its=iF.Restrict(dasl);

  for (var y=1;y<=its.Count;y++)
	{
	var Ap=its.Item(y);
	var ce=new ticklerObject(Ap);
		if (Ap!=null)
		{
		  if (fun>0){
        var ce=new ticklerObject(Ap);
    		var newRec=new ticklerRecord(ce);
    		store.add(newRec);
    		addDuration(newRec);
    		status=txtStatusGetting+"..." + Ap;}
    counter++;
		}
	}
store.sort("due","ASC");
try{var g=Ext.getCmp("tgrid");g.show();}catch(e){}
//status=txtReady;
return counter;
}

function ticklerObject(eID)
{
//add tickler items (26.appointments and 48.tasks and 43 messages)


try{this.entryID=eID.EntryID;}catch(e){return null;}
this.subject=eID.Subject;
this.contacts="";

this.importance=getJPriorityProperty(eID);
if (this.subject==null || this.subject==""){this.subject=txtNoSubject;}
this.attachment=null;
var t=new Date(eID.CreationTime);
this.created=t;
this.iclass=eID.Class;
this.sensitivity=eID.Sensitivity;
this.reminder=null;

	if (eID.Class==48 || eID.Class == 43)
	{//due date for tasks
    var mid=0;var did=0;
    	if( eID.Class == 43)
		var t = new Date(getJDueProperty(eID));
	else
    try{var t=new Date(eID.itemProperties.item(dueProperty).Value);}catch(e){}
    var dd=DisplayDate(t);if(dd==null || dd==""){dd="&nbsp;";}
    this.due=t;
    this.start=t;
    mid=(t.getMonth()+1);if (mid<10){mid="0"+mid;}
    did=t.getDate();if (did<10){did="0"+did;}
        if( eID.Class == 43)
    		this.status = "Active";
    	else{
        this.status=eID.Status;
        this.done=eID.Complete;
    	     }
    this.end=t;
    this.allday=false;
    this.location="";
    this.icon=this.iclass+";"+this.status+";"+this.sensitivity;
    this.duration="";
    	if (OLversion>11)
		  {
      //totalwork property for outlook 2007 or higher
      this.duration=eID.TotalWork;
      }
    this.done=eID.Complete;
    if (eID.ReminderSet==true){this.reminder=txtReminder+":"+eID.ReminderTime;}
    }
  else
  {//due dates for appointments
  try{var t=new Date(eID.Start);}catch(e){}
  var dd=DisplayDate(t);if(dd==null || dd==""){dd="&nbsp;";}
  this.due=t;
  this.start=eID.Start;
  try{var t=new Date(eID.End);}catch(e){}
  var dd=DisplayDate(t);if(dd==null || dd==""){dd="&nbsp;";}
  this.end=t;
  this.status="Active";
  this.allday=eID.AllDayEvent;
  this.icon=this.iclass+";0;"+this.sensitivity;
  this.location=eID.Location;
  this.duration=eID.Duration;
  this.done=eID.BillingInformation;
  if (eID.ReminderSet==true)
  {try{var rd=ol.Reminders.Item(this.subject);var nrd=new Date(rd.NextReminderDate);this.reminder=txtReminder+":"+DisplayDate(nrd)+"@"+DisplayTime(nrd);}catch(e){this.reminder=txtNoRemindata;}}
  }

  this.daypos=getDayPosition(this.due);
    if( eID.Class == 43)
	  this.isrecurring = false;
  else
  this.isrecurring=eID.IsRecurring;

	//links
		try{
    if (eID.Links.Count>0)
    {
		for (var x=1;x<=eID.Links.Count;x++)
		{this.contacts+=eID.Links.Item(x);}
		} }catch(e){}

	try{this.attachment=eID.UserProperties.Item("OLID").value;}catch(e){}
	try{this.tags=eID.itemProperties.Item(catProperty).Value;}catch(e){}
}

function renderTickSubject(v,m,r)
{
//render the subject line in grid
var tg=r.get("tags");
var id=r.get("entryID");
var st=r.get("status");
var cl=r.get("iclass");
var un=r.get("unread");
var rcr=r.get("isrecurring");
var con=r.get("contacts");
var rmr=r.get("reminder");
var don=r.get("done");
if (cl>99){return "<b>"+Ext.util.Format.htmlEncode(v)+"</b>";}
var rt="";
//var it=NSpace.GetItemFromID(id);
//try{
var tList=tagList('formtagdisplay',lastContext,tg, id);
//}catch(e){}
  var subj=Ext.util.Format.htmlEncode(v);
  if (st==2 || don=="[done]"){subj="<strike>"+subj+"</strike>";}
  if (un==true){subj="<b>"+subj+"</b>";}
  if (tList[1]==true){subj="<b>"+subj+"</b>";}
  var conicon="";
  var con2=con.replace(new RegExp("'","g")," ");
  if (notEmpty(con)){conicon="<a class=jellolink title='"+con2+"'><img src="+imgPath+"user.gif></a>";}
  if (cl==48){rt+=NextActionToggleIcons(tg,id);}
  if (notEmpty(rmr)){rt+="<img src=img\\clock.gif title='"+rmr+"'>&nbsp;";}
  if (rcr==true){rt+="<img src=img\\recur.gif>&nbsp;";}
  rt+=conicon+"&nbsp;"+subj+"&nbsp;"+tList[0];
 return rt;
}

function openTicklerItem(iid,otherGrid, oneitem)
{//open tickler item for editing

	if (notEmpty(iid))
	{
	var id=iid;
	var it=NSpace.GetItemFromID(id);
		if (it.Class==26)
		{var stt=it.Start;var dtt=new Date(it.Start);}
		else
		{var stt=new Date(it.itemProperties.item(dueProperty).Value);var dtt=stt;}
		if (it.Class==48){scAction(it.EntryID);}
	}
	else
	{

   var r;
  if( typeof(oneitem) != "undefined" && oneitem != null)
  	r = oneitem[0];
	else{
		if (notEmpty(otherGrid)){var g=otherGrid.getSelectionModel();}
  		else
  		var g=Ext.getCmp("tgrid").getSelectionModel();
	 r=g.getSelected();
	}
	var id=r.get("entryID");
	var stt=r.get("start");
	var dt=r.get("due");
	var it=NSpace.GetItemFromID(id);
	if (it.Class==48){var t=new Array(r);editAction(t);}
	}



  //open an appointment item
      if (it.Class==26 && it.isRecurring)
      {
          var msg="";
              try{
              var rc=it.GetRecurrencePattern.GetOccurrence(stt);
              msg=txtTikAppOccPrpt;
              var buts={yes:txtTikOpOcc,no:txtTikOpMas,cancel:txtCancel};
              }
              catch(e)
              {msg=txtTikOLError+"<BR><BR>"+e.description;
              var buts={no:txtTikOpMas,cancel:txtCancel};}

          Ext.Msg.show({
           title:txtAppointment,
           msg: msg,
           buttons: buts,
           fn: function(b,t){showApp(b,rc,it);},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });

          }else
          {if (it.Class==43 || it.Class==26 ){it.Display();}}


}

function showApp(b,rc,it)
{
//open apointment based on message box choice
    if (b=="no"){it.Display();}
    if (b=="yes"){rc.Display();}
}


function ticklerToolbar()
{
//render toolbar
    var menu2 = new Ext.menu.Menu({
        id: 'exportmenu',
        items: [{
                text: txtMnuSendPrw,
                tooltip:txtMnuSendPrwInfo,
                icon: 'img\\mail.gif',
                handler:function(){
        printCalendar(false);}
        },{
                text: txtPrintDirectly,
                tooltip:txtPrintDirectlyInfo,
                icon: 'img\\print.gif',
                handler:function(){
        printCalendar(true);}
        }


        ]}
    );

    var dateMenu = new Ext.menu.DateMenu({
        handler : function(dp, date){
        var a=new Date(date);var b=new Date(date);
		getTicklerFile(a,b,1,Ext.getCmp("tgrid").getStore(),true);
		thisRangeStart=a;thisRangeEnd=b;
		updateTicklerTitle();
        }
    });


    tbar1 = new Ext.Toolbar();
        tbar1.add({
        icon: 'img\\appoint.gif',
        cls:'x-btn-icon',
        id:'vnew',
        tooltip: '<b>'+txtCreate+'</b><br>'+txtTikCrNew+' (Insert)',
        handler : ticklerSelected
        },
        '-',{
        icon: 'img\\page_edit.gif',
        cls:'x-btn-icon',
        tooltip: '<b>Edit</b><br/>Edit selected item(s)',
        id:'vedit',
        handler : ticklerSelected
        },{
        icon: 'img\\page_delete.gif',
        cls:'x-btn-icon',
        id:'vdel',
        tooltip: txtDelItmInfo,
        handler : ticklerSelected
        },'-',{
        icon: 'img\\check.gif',
        cls:'x-btn-icon',
        id:'vdone',
        tooltip: txtCompleteItmInfo,
        handler : ticklerSelected
        },{
            tooltip: txtOpenInOutlook + '(Ctrl+Q)',
            cls:'x-btn-icon',
            id:'aoutlook',
            icon: 'img\\info.gif',
             listeners:{click: function(b,e){ticklerInOutlook();} }
        },'-',{
            icon: 'img\\calendar.gif',
            cls:'x-btn-icon',
            id:'popdate',
            tooltip: txtTikGoDate+' [Ctrl+D]',
            menu: dateMenu
            },{
                text: txtDay,
                id:'vday',
                enableToggle:true,
                toggleGroup:'ctype',
                handler: ticklerRangeSelected
        },{
                text: txtWeek,
                id:'vweek',
                enableToggle:true,
                toggleGroup:'ctype',
                handler: ticklerRangeSelected
        },{
                text: txtMonth,
                id:'vmonth',
                enableToggle:true,
                toggleGroup:'ctype',
                handler: ticklerRangeSelected
        },{
                text: txtAgenda,
                id:'vagenda',
                enableToggle:true,
                toggleGroup:'ctype',
                handler: ticklerRangeSelected
        },'-',{
        iconCls: 'x-tbar-page-prev',
        tooltip: txtTikPrvRng,
        id:'vprev',
			  handler:ticklerRangeSelected
        },{
        iconCls: 'x-tbar-page-next',
        tooltip: txtTikNexRng,
        id:'vnext',
        handler:ticklerRangeSelected
        },'->',{
        icon: 'img\\refresh.gif',
        cls:'x-btn-icon',
        tooltip: txtRefreshInfo,
        id:'vrefresh',
		handler:ticklerSelected
        },{
            text:txtMenuOtherFunctions,
            menu: menu2
        }

         );
return tbar1;
}

function DisplayDuration(v,i,m)
{
//render duration field
v=Math.round(v);
if (v==0){return null;}
if (v<60){v=v+" "+txtMins;}
if (v>60){v=(v/60)+" "+txtHours;}
if (v==60){v="1 "+txtHour;}
if (v==1440){v="1 "+txtDay;}
if (v>1440){v=(v/1440)+" "+txtDays;}
if (v==10080){v="1 "+txtWeek;}
if (v>10080){v=(v/10080)+" "+txtWeeks;}
return v;
}

function ticklerSelected(btn,vl, oneitem)
{
  if (btn.id=="vnew")
  {//new appointment
  var iF=NSpace.GetFolderFromID(jello.calendarFolder).Items;
  var napp=iF.Add();
  napp.Display();
  }

  if (btn.id=="vedit" || btn.id=="avedit")
  {//edit selected item
  openTicklerItem(null,null,oneitem);
  }

  if (btn.id=="vdel" || btn.id=="vdone" || btn.id=="avdel" || btn.id=="avdone")
  {//delete selected item(s)
  loopSelectedTicklers(btn.id, oneitem);
  }

  if (btn.id=="vrefresh")
  {//refresh view
	   try{
     var g=getActiveGrid();
     if (g.getId()=="tgrid")
     {getTicklerFile(thisRangeStart,thisRangeEnd,1,g.getStore(),true);}
    }catch(e){}

  }

}

function loopSelectedTicklers(act, oneitem)
{//loop selected ticklers (tasks/appts) and perform an action (delete, done)

if (act=="vdel" || act=="avdel"){var amsg=txtMsgDelItem;
var co=confirm(amsg+"\n"+txtActOnOccMsg);
if (co==false){return;}
}
var cit;
if( typeof(oneitem) != "undefined" && oneitem != null)
	cit = oneitem;
else
	cit=getCheckedItems(thisGrid);
if (cit==false){return;}
var cl=cit.length;
if (cl==0){return;}

		for (var x=0;x<cl;x++)
		{
		var id=cit[x].get("entryID");
		var it=NSpace.GetItemFromID(id);

		if (it.Class==26)
		{//appointment (check for occurrence. if exists use occurrence)
			if(it.IsRecurring)
			{
			var dt=cit[x].get("due");
			var stt=cit[x].get("start");
			try{var rc=it.GetRecurrencePattern.GetOccurrence(stt);it=rc;}catch(e){}
			}
		}

			if (act=="vdel" || act=="avdel")
			{//delete
			updateRecordItem(cit[x],true,it);
			it.Delete();
			}

			if (act=="vdone" || act=="avdone")
			{//mark completed
			completeTickler(it);
			if (jello.CalendarShowCompleted=="0" || jello.CalendarShowCompleted==false )
			{updateRecordItem(cit[x],true,it);}else{updateRecordItem(cit[x],false,it);}
			}
		}
}

function completeTickler(it)
{//mark tickler as complete based on item's type
	if (it.Class==26)
	{//appointment
	var apv=it.BillingInformation;
	if (apv=="[done]"){apv="";}else{apv="[done]";}
	it.BillingInformation=apv;
	it.Save();
	}
	else
	{//task
	var id=it.EntryID;
	actionDone(id);
	}
}

function ticklerRangeSelected(btn,vl)
{
//tickler toolbar actions management

var isMove=false;
if (btn.id!="vnext" && btn.id!="vprev")
{thisRangeStart=new Date();thisRangeEnd=new Date();}

if (btn.id=="vday"){
thisCalView="vday";
}

if (btn.id=="vweek"){
thisCalView="vweek";

var now =  new  Date();
var nowDayOfWeek = now.getDay()-1;
var nowDay = now.getDate();
var nowMonth = now.getMonth();
var nowYear = now.getYear();
nowYear += (nowYear < 2000) ? 1900 : 0;
thisRangeStart = new Date(nowYear, nowMonth, nowDay - nowDayOfWeek);
thisRangeEnd = new Date(nowYear, nowMonth, nowDay + (6 - nowDayOfWeek));
}

if (btn.id=="vmonth"){
thisCalView="vmonth";
thisRangeStart.setDate(1);
var ldom=daysInMonth(thisRangeStart.getMonth()+1,thisRangeStart.getFullYear());
thisRangeEnd.setDate(ldom);
}

if (btn.id=="vagenda"){
thisCalView="vagenda";
//thisRangeEnd.setDate(thisRangeStart.getDate()+jello.agendaDays);
thisRangeEnd.setTime(thisRangeStart.getTime() + (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));
}

if (btn.id=="vnext")
{//next range
  isMove=true;
  if(thisCalView=="vday"){thisRangeStart.setDate((thisRangeStart.getDate()+1));thisRangeEnd=thisRangeStart.clone();}
  if(thisCalView=="vweek"){thisRangeStart.setDate((thisRangeStart.getDate()+7));thisRangeEnd.setDate((thisRangeEnd.getDate()+7));}
  if(thisCalView=="vmonth"){thisRangeStart.setMonth((thisRangeStart.getMonth())+1);
  thisRangeEnd=thisRangeStart.clone();
  var ldom=daysInMonth(thisRangeEnd.getMonth()+1,thisRangeEnd.getFullYear());
  thisRangeEnd.setDate(ldom);
  }
    if(thisCalView=="vagenda")
    {
    //thisRangeStart.setDate((thisRangeStart.getDate()+jello.agendaDays));thisRangeEnd.setDate((thisRangeEnd.getDate()+jello.agendaDays));
    thisRangeStart.setTime(thisRangeStart.getTime() + (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));
    thisRangeEnd.setTime(thisRangeEnd.getTime() + (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));
    }
}

if (btn.id=="vprev")
{//previous range
  isMove=true;
  if(thisCalView=="vday"){thisRangeStart.setDate((thisRangeStart.getDate()-1));thisRangeEnd=thisRangeStart.clone();}
  if(thisCalView=="vweek"){thisRangeStart.setDate((thisRangeStart.getDate()-7));thisRangeEnd.setDate((thisRangeEnd.getDate()-7));}
  if(thisCalView=="vmonth"){thisRangeStart.setMonth((thisRangeStart.getMonth())-1);
  thisRangeEnd=thisRangeStart.clone();
  var ldom=daysInMonth(thisRangeEnd.getMonth()+1,thisRangeEnd.getFullYear());
  thisRangeEnd.setDate(ldom);
  }
    if(thisCalView=="vagenda")
    {
//    thisRangeStart.setDate((thisRangeStart.getDate()-(jello.agendaDays)));thisRangeEnd.setDate((thisRangeEnd.getDate()-(jello.agendaDays)));
    thisRangeStart.setTime(thisRangeStart.getTime() - (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));
    thisRangeEnd.setTime(thisRangeEnd.getTime() - (24 * 60 * 60 * 1000 * parseInt(jello.agendaDays, 10)));

    }
}

	var store=Ext.getCmp("tgrid").getStore();
	getTicklerFile(thisRangeStart,thisRangeEnd,1,store,true,isMove);
	updateTicklerTitle();
	updateAFooterCounter(store);
try{
	Ext.getCmp("tgrid").getView().focusRow(0);
	}catch(e){}
}

function updateTicklerTitle(noUpdate)
{//update screen title
    var ttl=DisplayDate(thisRangeStart)+" - "+DisplayDate(thisRangeEnd);
	var ddf=dayDifSimple(thisRangeEnd,thisRangeStart);
    if (ddf==1){ttl=" ["+thisRangeStart.toLocaleDateString()+"]";}
    if (ddf==7){ttl+=" ("+txtWeek+" "+thisRangeStart.format('W')+")";}
    if (ddf>27 && ddf<32){ttl=thisRangeStart.format('F o');}
    if (noUpdate==true)
    {return ttl;}
    else
    {Ext.getCmp("portalpanel").setTitle("<img src=img\\ticklers16.png style=float:left;> "+txtTicklers+"  <span class=ticklertitle>"+ttl+showTotalDuration()+"</span>");}
}

function DisplayCalDate(t,m,r)
{
//display of date and time for calendar views
var sd=new Date(r.get("start"));
var tmm=sd.format("H:i");
if (jello.timeFormat==1 || jello.timeFormat=="1"){tmm=sd.format("g:i");}
var ret=DisplayAppDate(sd);
var ad=r.get("allday");
if (ad==false){ret+=" <span class=caltime>["+tmm+"]</span>";}
else{ret+=" <span class=callday>["+txtAllDay+"]</span>";}
return ret;
}

function getDayPosition(due)
{
//return day position based on today
var td=new Date();
var duew=parseInt(due.format('W'));
var tdw=parseInt(td.format('W'));
td.setHours(00,00,00);
due.setHours(00,00,00);
var ret="";
var df=dayDifSimple(due,td);
if (df>1){ret="<span class=9futuregroup>"+txtFuture+"</span>";}
if (df<0){ret="<span class=1pastgroup>"+txtPastDue+"</span>";}
if (df>1 && duew==tdw){ret="<span class=3latergroup>"+txtLaterWeek+"</span>";}
if (df>1 && duew==(tdw+1)){ret="<span class=4nextwgroup>"+txtNextWeek+"</span>";}
if (df==0){ret="<span class=2todaygroup>"+txtToday+"</span>";}
if (df==1){ret="<span class=2todaygroup>"+txtTomorrow+"</span>";}
return ret;
}

function printCalendar(direct)
{
//print calendar grid
var store=Ext.getCmp("tgrid").getStore();
var counter=0;
var ret="<head><style>.printChars{height:30px;font-size: 14px;font-family: 'Segoe UI', Verdana, 'Trebuchet MS', Arial, Sans-serif;border-bottom: 1px solid Gainsboro;}</style></head><body><table width=100% cellpadding=0 cellspacing=0>";
ret+="<tr><td class=printChars><b>"+txtTickler+"</b></td><td class=printChars><b>"+txtStartDate+"</b></td><td class=printChars><b>"+txtDuration+"</b></td><td class=printChars><b>"+txtLocation+"</b></td></tr>";

  var toPrint = function(rec){
  ret+="<tr><td class=printChars><span style=font-family:Wingdings;font-size:16px;>q</span> "+rec.get("subject")+"</td><td class=printChars>"+DisplayCalDate(rec.get("due"),null,rec)+"</td><td class=printChars>"+DisplayDuration(rec.get("duration"))+"</td><td class=printChars>"+rec.get("location")+"</td></tr>";
  counter++;
  };

store.each(toPrint);
ret+="</table><br><br><span class=printChars>"+counter+" "+txtItemItems+"</span>";


var ff=NSpace.GetFolderFromID(jello.inboxFolder).Items;
var it=ff.Add(6);
it.Subject="Ticklers: "+updateTicklerTitle(true);
it.HTMLBody=ret;
if (direct==true){it.PrintOut();Ext.info.msg(txtPrinting,txtPrintOK);}
else{it.Display();}

}


function ticklerInOutlook(oneitem)
{
var gg=getActiveGrid();
var g=gg.getSelectionModel();
var r;
if( typeof(oneitem) != "undefined" && oneitem != null)
	r = oneitem[0];
else
	r=g.getSelected();
try{var id=r.get("entryID");
var it=NSpace.GetItemFromID(id);}catch(e){alert(txtInvalid);return;}
if (it.Class==48){it.Display();return;}
if( typeof(oneitem) != "undefined" && oneitem != null)
	openTicklerItem(null,null,oneitem);
else
	openTicklerItem();
}

function ticklerContextMenu(rec, id, it, inlist)
{
try{Ext.getCmp("atrclickmenu").destroy();}catch(e){}
	var oneitem= new Array(); oneitem.push(rec);
	return new Ext.menu.Menu({
		id: 'atrclickmenu',
		items: [
		{
		text: txtBody,
		id:'avmbody',
		handler:function(){
		popTaskBody(id,rec);}
		},
		{
        icon: 'img\\page_edit.gif',
        text: txtEdit,
        id:'avedit',
       listeners:{click: function(b,e){ticklerClickForwarder(b,e,inlist,rec);} }
        },
      {
        icon: 'img\\page_delete.gif',
        id:'avdel',
        text: txtDelItmInfo,
        listeners:{click: function(b,e){ticklerClickForwarder(b,e,inlist,rec);} }
        },
		{
        icon: 'img\\check.gif',
        id:'avdone',
        text: txtCompleteItmInfo,
        listeners:{click: function(b,e){ticklerClickForwarder(b,e,inlist,rec);} }
        },
		{
			text: txtATaskfrom,
			handler:function(){
				createTaskFromAppointment(it);}
		},
		{
            text: txtOpenInOutlook,
            id:'actoutlook',
            icon: 'img\\info.gif',
            udata: (inlist? null:oneitem),
            listeners:{click: function(b,e){ticklerInOutlook(b.udata);} }
        },
		{
			text: txtShowWithTag,
			id: 'avsametasks',
			icon: 'img\\task.gif',
			listeners: {click: function(b,e){tasksWithTag(rec);} }
		}

		]
	});
}

function ticklerClickForwarder(b,e,inlist, rec)
{
	if( inlist){
			ticklerSelected(b,e);
	}else{
		var x = new Array();
		x.push(rec);
		ticklerSelected(b,null,x);
	}
}

function createTaskFromAppointment(it)
{
try{
	var newA = NSpace.GetFolderFromID(jello.calendarFolder).Items.Add();
}catch(e){return;}
try{
	newA.Subject = txtAbout +" - " + it.Subject;
	newA.Body = it.Body;
	newA.Categories = it.Categories;
	setAttachmentProperty(newA,it.EntryID,true);
	newA.Save();
	newA.Display();
}catch(e){newA.Delete();}
}

function showTotalDuration()
{
//render te total duration of ticklers list
var du=totalDuration;
var ret="&nbsp<span class=caltime><font size=1>(Total:"+du+" minutes)</font></span>";
return ret;
}

function addDuration(newRec)
{//add duration of item to total duration. Exclude allday items
if (newRec.data.allday==false){totalDuration+=newRec.data.duration;}
}