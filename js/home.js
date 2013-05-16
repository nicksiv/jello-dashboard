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

var customizingWidget=null;
var ImportantPastDueCount=0;

//define tools
var tSet={
        id:'gear',
        handler: function(e, target, panel){
        var pid=panel.getId();
        pid=pid.replace("w","");
        editWidget(pid); 
        }
    };
var tClose={
        id:'close',
        handler: function(e, target, panel){
        var pid=panel.getId();
        pid=pid.replace("w","");
        var question=removeWidget(pid);
        if(question){panel.ownerCt.remove(panel, true);}
        }
    };
    

var tCol={
        id:'save',
        text:'Collect',
        tooltip:'Collect',
        handler: function(e, target, panel){
        var pid=panel.getId();
        pid=pid.replace("w","");
		var coltxt=Ext.getCmp("wgrid:"+pid).getRawValue();
        
            var c=importText(true,coltxt);
            if (c==true){importText(false,coltxt);return;}
        
        }
    };

var tClear={
        id:'restore',
        text:'Clear',
        tooltip:'Clear',
        handler: function(e, target, panel){
        
        var choice=confirm("Clear all text from the Postit widget?");
        if (choice){
        var pid=panel.getId();
        pid=pid.replace("w","");
		    Ext.getCmp("wgrid:"+pid).setValue('');
        wPostItSave(Ext.getCmp("wgrid:"+pid),pid)
         }
        }
    };
    
    var tRefresh={
        id:'refresh',
        handler: function(e, target, panel){
        var pid=panel.getId();
        pid=pid.replace("w","");
        doWidgetRefresh(panel,pid);
        
		}
    };

    
//define widgets datablock
var widgetStore=new Ext.data.JsonStore({
        fields: [
           {name: 'type'},
           {name: 'column',type:'int'},
           {name: 'row',type:'int'},
           {name: 'autorefresh',type:'boolean'},
           {name: 'settings'},
           {name: 'widgetID',type: 'int'},
           {name: 'collapsed',type: 'boolean'},
           {name: 'height',type: 'int'},
           {name: 'content'}
        ],
        data:jello.widgets
    });
    
 var widgetRecord = Ext.data.Record.create([
           {name: 'type'},
           {name: 'column',type: 'int'},
           {name: 'row',type: 'int'},
           {name: 'autorefresh',type: 'boolean'},
           {name: 'settings'},
           {name: 'widgetID',type: 'int'},
           {name: 'collapsed',type: 'boolean'},
           {name: 'height',type: 'int'},
           {name: 'content'}
	]);



Ext.ux.Portal = Ext.extend(Ext.Panel, {
    layout: 'column',
    autoScroll:true,
    cls:'x-portal',
    defaultType: 'portalcolumn',
    
    initComponent : function(){
        Ext.ux.Portal.superclass.initComponent.call(this);
        this.addEvents({
            validatedrop:true,
            beforedragover:true,
            dragover:true,
            beforedrop:true,
            drop:true
        });
    },

    initEvents : function(){
        Ext.ux.Portal.superclass.initEvents.call(this);
        this.dd = new Ext.ux.Portal.DropZone(this, this.dropConfig);
    },
    
    beforeDestroy: function() {
        if(this.dd){
            this.dd.unreg();
        }
        Ext.ux.Portal.superclass.beforeDestroy.call(this);
    }
});
Ext.reg('portal', Ext.ux.Portal);

Ext.ux.Portal.DropZone = function(portal, cfg){
    this.portal = portal;
    Ext.dd.ScrollManager.register(portal.body);
    Ext.ux.Portal.DropZone.superclass.constructor.call(this, portal.bwrap.dom, cfg);
    portal.body.ddScrollConfig = this.ddScrollConfig;
};


Ext.extend(Ext.ux.Portal.DropZone, Ext.dd.DropTarget, {
    ddScrollConfig : {
        vthresh: 50,
        hthresh: -1,
        animate: true,
        increment: 200
    },

    createEvent : function(dd, e, data, col, c, pos){
        return {
            portal: this.portal,
            panel: data.panel,
            columnIndex: col,
            column: c,
            position: pos,
            data: data,
            source: dd,
            rawEvent: e,
            status: this.dropAllowed
        };
    },

    notifyOver : function(dd, e, data){
        var xy = e.getXY(), portal = this.portal, px = dd.proxy;

        // case column widths
        if(!this.grid){
            this.grid = this.getGrid();
        }

        // handle case scroll where scrollbars appear during drag
        var cw = portal.body.dom.clientWidth;
        if(!this.lastCW){
            this.lastCW = cw;
        }else if(this.lastCW != cw){
            this.lastCW = cw;
            portal.doLayout();
            this.grid = this.getGrid();
        }

        // determine column
        var col = 0, xs = this.grid.columnX, cmatch = false;
        for(var len = xs.length; col < len; col++){
            if(xy[0] < (xs[col].x + xs[col].w)){
                cmatch = true;
                break;
            }
        }
        // no match, fix last index
        if(!cmatch){
            col--;
        }

        // find insert position
        var p, match = false, pos = 0,
            c = portal.items.itemAt(col),
            items = c.items.items;

        for(var len = items.length; pos < len; pos++){
            p = items[pos];
            var h = p.el.getHeight();
            if(h !== 0 && (p.el.getY()+(h/2)) > xy[1]){
                match = true;
                break;
            }
        }

        var overEvent = this.createEvent(dd, e, data, col, c,
                match && p ? pos : c.items.getCount());

        if(portal.fireEvent('validatedrop', overEvent) !== false &&
           portal.fireEvent('beforedragover', overEvent) !== false){

            // make sure proxy width is fluid
            px.getProxy().setWidth('auto');

            if(p){
                px.moveProxy(p.el.dom.parentNode, match ? p.el.dom : null);
            }else{
                px.moveProxy(c.el.dom, null);
            }

            this.lastPos = {c: c, col: col, p: match && p ? pos : false};
            this.scrollPos = portal.body.getScroll();

            portal.fireEvent('dragover', overEvent);

            return overEvent.status;;
        }else{
            return overEvent.status;
        }

    },

    notifyOut : function(){
        delete this.grid;
    },

    notifyDrop : function(dd, e, data){
        delete this.grid;
        if(!this.lastPos){
            return;
        }
        var c = this.lastPos.c, col = this.lastPos.col, pos = this.lastPos.p;

        var dropEvent = this.createEvent(dd, e, data, col, c,
                pos !== false ? pos : c.items.getCount());

        if(this.portal.fireEvent('validatedrop', dropEvent) !== false &&
           this.portal.fireEvent('beforedrop', dropEvent) !== false){

            dd.proxy.getProxy().remove();
            dd.panel.el.dom.parentNode.removeChild(dd.panel.el.dom);
            if(pos !== false){
                c.insert(pos, dd.panel);
            }else{
                c.add(dd.panel);
            }
            
            c.doLayout();

            this.portal.fireEvent('drop', dropEvent);

            // scroll position is lost on drop, fix it
            var st = this.scrollPos.top;
            if(st){
                var d = this.portal.body.dom;
                setTimeout(function(){
                    d.scrollTop = st;
                }, 10);
            }

        }
        delete this.lastPos;
        updateWidgetOrder();
    },

    // internal cache of body and column coords
    getGrid : function(){
        var box = this.portal.bwrap.getBox();
        box.columnX = [];
        this.portal.items.each(function(c){
             box.columnX.push({x: c.el.getX(), w: c.el.getWidth()});
        });
        return box;
    },

    // unregister the dropzone from ScrollManager
    unreg: function() {
        try{Ext.dd.ScrollManager.unregister(this.portal.body);}catch(e){}
        Ext.ux.Portal.DropZone.superclass.unreg.call(this);
    }
});


Ext.ux.Pcolumn = Ext.extend(Ext.layout.ColumnLayout, {
    scrollOffset:18
});




Ext.ux.PortalColumn = Ext.extend(Ext.Container, {
    layout: 'anchor',
    autoEl: 'div',
    defaultType: 'portlet',
    cls:'x-portal-column'
});
Ext.reg('portalcolumn', Ext.ux.PortalColumn);


Ext.ux.Portlet = Ext.extend(Ext.Panel, {
    anchor: '100%',
    frame:true,
    collapsible:true,
    autoWidth:true,
    draggable:true,
    cls:'x-portlet'
});
Ext.reg('portlet', Ext.ux.Portlet);



function pHome()
{
	if(pHome.arguments.length == 0)
	{
  initScreen(false,"pHome()");
	try{Ext.getCmp("homepop").destroy();}catch(e){}
	}
	
checkWidgets();        
updateCounter(-1);
Ext.getCmp("previewpanel").update(txtWdImpWel+" "+jelloVersion);
var wdth=(1/jello.widgetColumns);
//wdth=wdth-0.008;
var col="[";
var rp="0px";
//create the columns
//try{
	for (var x=0;x<jello.widgetColumns;x++)
	{
	if (x==(jello.widgetColumns-1)){rp="19px";}
	var cid="widgetColumn"+x;
	col+="{columnWidth:"+wdth+",id:'col"+x+"',style:'padding:3px "+rp+" 3px 3px',id:'"+cid+"'";
	var theItems=renderWidgets(x);
  if (theItems!="00000"){col+=",items:["+theItems+"]";}
  col+="},";
	}

	col=col.substr(0,(col.length-1));
	col+="]";
	
if(pHome.arguments.length == 0){    
	      /*var hm=new Ext.ux.Portal ({
	            renderTo: 'main',
	            xtype:'portal',
	            id:'portalpanel',
	            margins:'2 2 2 0',
	            height:(panelHeight-30),
	            width:(panelWidth-5),
	            items:eval(col)
	            
	    }); */
	 var ppnl=Ext.getCmp("portalpanel");
   
   
   ppnl.add(eval(col));
   ppnl.setAutoScroll(true);
///*************************************
   ppnl.doLayout();

 	 ppnl.setTitle("<img src=img\\home16.png style=float:left;>"+txtHome+"&nbsp;&nbsp;&nbsp;&nbsp;<a class=jellolink style=font-weight:normal; onclick=addWidgetScreen()>"+txtAddWidget+"</a>&nbsp;|&nbsp;<a class=jellolink style=font-weight:normal; onclick=customizeWidgetScreen()>"+txtColumns+"</a>&nbsp;|&nbsp;<a class=jellolink style=font-weight:normal; onclick=resetWidgets()>"+txtReset+" Widgets</a>");
	 
  }else
		return new Ext.ux.Portal ({
	            xtype:'portal',
	            id:'portalpop',
	            margins:'2 2 2 0',
	            height:(panelHeight-30),
	            width:(panelWidth-5),
	            items:eval(col)
	            
	    });
	    
/*
}catch(e)

	{
	//widgets error
		var msg="<b>"+e+"</b><br>"+txtDashError;
		var buts={yes:txtYes,no:txtNo};
          Ext.Msg.show({
           title:txtDashErrorTtl,
           msg: msg,
           buttons: buts,
           fn: function(b,t){if (b=="yes"){resetWidgets();}},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });	
	}
	*/
	if(pHome.arguments.length != 0) 
		return null;
//	homie.style.display="none";
	firstLoadHome=false;
}

function aboutStuff()
{
var ret="<div style=padding-left:0px><h1><p align=center><img src=img\\biglogo.gif></p><p align=center>"+jelloVersion+"</h1></p><br>";
ret+="<p align=center>By Nicolas Sivridis 2006-2010<br>";
ret+="<img src=img\\gplv3.png> Licensed under Open Source Lisence GPLv3 <i>(see enclosed copying.txt file for details)</i><br><br>";
ret+="<img src=img\\extjs.gif> Built with the open source <a href=http://extjs.com><b>Extjs</b></a> development libraries<br>";
ret+="<br>Thanks to <a href=http://www.famfamfam.com/><b>famfamfam</b></a> for the Open source Icon Set <a href=http://www.famfamfam.com/lab/icons/silk/><b>Silk</b></a><br>";
ret+="<br> Compressed some code using the <b>http://javascriptcompressor.com</b>";
ret+="<br>Thanks to <b>Phil Shevrin</b> for his code fragments<br>";
ret+="<br>Current Language Translation: <b>"+txtLangCredits+"</b>";
ret+="<br><br>More information at the <a class=jellolinktop href=http://jello-dashboard.net>Jello Dashboard website</a> and the <a class=jellolinktop href=http://jello-dashboard.net/forum>Jello Dashboard forum</a>";
ret+="<br><span style='text-decoration:underline;color:red;'>You can support this project and help it grow by contributing a small Donation</span><br><br><a href='https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=druqbar%40gmail%2ecom&item_name=Jello%2eDashboard&no_shipping=2&no_note=1&tax=0&currency_code=EUR&bn=PP%2dDonationsBF&charset=UTF%2d8'><img src=img/donateCC_LG_global.gif></a><br><br><br>";
ret+="</p></div>";
return ret;
}

function about()
{
//about jello screen
//initScreen(false,"about()");
showOLViewCtl(false);
var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
    //    title: txtTtlLatestAt,
        height:750,
        bodyStyle:'padding:5px 5px 0 5px',
        floating:false,
        labelWidth:480,
        id:'latestform',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',
        autoScroll:true,
       	html:aboutStuff()
       

    });

    var win = new Ext.Window({
        title: jelloVersion,
        width: 750,
        height:500,
        id:'aboutform',
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
        listeners: {destroy:function(t){}},
        buttons: [{
            text: txtClose,
            tooltip:txtClose+' [F12]',
            listeners:{
            click: function(b,e){
            win.destroy();showOLViewCtl(true);}
            }
        }]
    });

    win.show();
    win.setActive(true);
}

function jelloDocs()
{
//create/open a special journal entry containing help stuff
  var jfldr=draftItems;
  var je=jfldr.Items;
	var j=je.Restrict("[Subject]='Jello.Help'");
		if (j.Count==0)
		{
		status=txtMsgWait+"...";
		var ja=je.Add(6);
		ja.subject="Jello.Help";
    var uri=getAppPath();
		ja.HTMLBody="<h2>Jello Resources</h2>Open the attached files for additional information on the project.<br>Remember to read the Documentation pdf.<br><br><a href=http://jello-dashboard.net>Jello.Dashboard home page</a><br><a href='http://jello-dashboard.net/forum' target='_blank'>Jello.Forum</a><br><a href=https://sourceforge.net/projects/jello5/>Jello.Dashboard open source project at Sourceforge</a><br><h3>For questions, bugs and feature requests visit the <a href=http://jello-dashboard.net/forum>Jello forum</a></h3><br>"+txtHlpInfo;
		if (uri.substr(0,5)=="file:"){uri=uri.replace("file:","");}
		//ja.Attachments.Add(uri+"jello5-Help.pdf");
		ja.Attachments.Add(uri+"copying.txt");
		ja.Attachments.Add(uri+"readme.txt");
		ja.Post();
		alert(txtHlpPostGen);
		}
	    else
	  {
    j(1).Display();
    }
	    
status="";
}

function renderWidgets(id)
{
//get user widgets. 
var ret="";
widgetStore.filter("column",new RegExp("^"+id+"$"));
	var c=widgetStore.getCount();
	if (c>0)	
	{
	//loop through widgets and add them
		var addWidget=function(rec){
		var wId=rec.get("type");
		  if (!notEmpty(wId) || wId=="undefined")
        {newW= "{title:'ERROR',collapsed:false, id:'w99999',html:'Widget Format Error due to application upgrade!<br><br><b>Please <a class=jellolink style=font-weight:normal; onclick=resetWidgets()>"+txtReset+" Widgets</a></b><br><br><a class=jellolinktop style=font-weight:normal;text-decoration:underline; onclick=resetWidgets(true)><b>["+txtReset+"]</b></a>'}";}
      else
      {
          		var wiID=rec.get("widgetID");
          		var iswcol=rec.get("collapsed");
          		if(!notEmpty(iswcol)){iswcol="false";}
          		var winfo=getWidgetInfo(wId);
          		var wtitle=winfo[0];
          		var tls="tools:[tRefresh,tSet,tClose],";
          		if (winfo[2]){tls="tools:[tSet],";}
          		if (winfo[3]){tls="tools:[tClose],";}
          		if (winfo[2] && winfo[3]){tls="tools:[tRefresh],";}
          		if (winfo[2] && winfo[3]==false){tls="tools:[tSet,tRefresh],";}
          		if (wtitle==txtWidTJPostit){tls="tools:[tClear,tCol,tSet,tClose],";}
          		//create the widget
          		if (jello.startup==3 && firstLoadHome){iswcol="true";}
			if (iswcol=="false")
			{
          		if(wtitle == txtTicklers)
          			{var content=eval(winfo[1].replace("()","("+wiID+",null,'')"));}
          		else
          			{var content=eval(winfo[1].replace("()","("+wiID+")"));}
          	}
          	else
          	{var content="";
          			if(wtitle == txtWidTJTagAct || wtitle==txtwOLSearch)
          			{eval(winfo[1].replace("()","("+wiID+",null,null,null,null)"));}
          			
          	}	
              var newW="{title:'"+wtitle+"',iconCls:'"+wId+"',collapsed:"+iswcol+", id:'w"+wiID+"',listeners:{resize:resizePortlet,collapse:function(p){colexpWidget(p,0);},expand:function(p){colexpWidget(p,1);}},"+tls+"items:["+content+"]}";	
          			try{var textW=Ext.util.JSON.decode(newW);}
          			catch(e){
          			if (tls.substr(tls.length-1,1)==","){tls=tls.substr(0,tls.length-1);}
          			newW="{title:'"+wtitle+"',collapsed:false, id:'w"+wiID+"',html:'Widget Error! <a class=jellolink style=font-weight:normal; onclick=resetWidgets()>'+txtReset+' Widgets</a>',"+tls+"}";
          			}
      }
		   //Outlook 2010 can have native views only in IE or hta mode
       if (OLversion>13.9 && conStatus=="Outlook" && wId=="OLW")
       {
       newW= "{title:'Outlook View Control',collapsed:false, id:'w"+wiID+"',html:'<br>Outlook 2010 does not allow the full use of the Outlook View Control inside Outlook. To use this widget run Jello Dashboard standalone or into Internet Explorer<br>&nbsp;',tools:[tClose]}";
       }


		ret+=newW+",";
		};		
		widgetStore.sort("row","ASC");
		widgetStore.each(addWidget);
		
	}
widgetStore.clearFilter();

var rl=ret.length;
if (rl>0){
ret=ret.substr(0,(ret.length-1));
}
if(!notEmpty(ret)){ret="00000";}
return ret;
}

function getWidgetInfo(id)
{
//return widget's title
	try{
	var a=new Array();
	var ix=widgetMapStore.find("id",new RegExp("^"+id+"$"));
	var r=widgetMapStore.getAt(ix);
	a.push(r.get("name"));
	a.push(r.get("function"));
	a.push(r.get("noremove"));
	a.push(r.get("nocustomize"));
	return a;
	}catch(e){return null;}
}

function checkWidgets()
{//check for widget records
status=txtMsgStChecking;
var lcol=1;
if (jello.widgetColumns==1){lcol=0;}
var c=widgetStore.getCount();

	if (c==0)
	{
	//If no widget exists create 2 default ones
	var tr=new widgetRecord({
    type:'WFL',
    column:0,
    row:0,
    widgetID:1
	});
	widgetStore.add(tr);
	var tr=new widgetRecord({
    type:'IMP',
    column:lcol,
    row:0,
	widgetID:2,
	settings:'Yes|||Yes'
	});
	widgetStore.add(tr);
	syncStore(widgetStore,"jello.widgets");
	//save settings
	jese.saveCurrent();
	}
status="";
}

function addWidgetScreen()
{
//show the add widget screen

var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: txtAddWidget,
        height:370,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'widgetform',
        //iconCls:'tagformicon',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',
        items: [
            new Ext.form.Label({
				cls:'formattlist',
			            height:250,
				html:widgetList()
            })
        ]

    });

    var win = new Ext.Window({
        title: 'Jello Home',
        width: 480,
        height:370,
        id:'thewidgetform',
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
            text: txtClose,
            tooltip:txtClose,
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

function widgetList()
{
//render a widget list
var ret="";
widgetMapStore.filter("noremove",false);
	var wrtWidg=function(rec){
	var torender=true;
		var wiID=rec.get("id");
		var wic=rec.get("icon");
		if (!notEmpty(wic)){wic="icon_extension.gif";}
		if(rec.get("unique") && widgetExists(wiID)){torender=false;}
		   //Outlook 2010 can have native views only in IE or hta mode
       if (OLversion>13 && conStatus=="Outlook" && wiID=="OLW"){torender=false;}

		if (torender)		
		{ret+="<a style='background-image:url("+imgPath+wic+");background-repeat:no-repeat;padding-left:18px;width:200;height:25;color:blue;' class=jellolink onclick=insertWidget('"+wiID+"'); title='"+txtWdgClickAdd+"'><b>"+rec.get("name")+"</b></a><br>"+rec.get("description")+"<hr>";}
	};

widgetMapStore.each(wrtWidg);
widgetMapStore.clearFilter();
return ret;
}

function widgetExists(id)
{
//check if widget already exists in users configuation
var exst=false;
var ix=widgetStore.find("type",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);
if (r!=null){exst=true;}
return exst;
}

function insertWidget(wid)
{
//insert a widget @ home
var ccount=jello.widgetColumns;
ccount--;
widgetStore.filter("column",ccount);
widgetStore.sort("row","DESC");
var wsr=widgetStore.getAt(0);
if(wsr!=null){var crow=wsr.get("row");}else{crow=-1;}
crow++;
widgetStore.clearFilter();
widgetStore.sort("widgetID","DESC");
var newWID=widgetStore.getAt(0).get("widgetID");
if (notEmpty(newWID)){newWID++;}else{newWID=1;}
var sets=getWidgetDefaultSettings(wid);
	var tr=new widgetRecord({
    type:wid,
    column:ccount,
    row:crow,
    autorefresh:false,
    widgetID:newWID,
    collapsed:false,
    height:150,
    settings:sets
	});
	widgetStore.add(tr);
	syncStore(widgetStore,"jello.widgets");
	jese.saveCurrent();
Ext.getCmp("thewidgetform").destroy();
pHome();
}

function getWidgetSettings(id)
{
//return array with specific widget's settings
var ws=null;
var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);
if (r!=null){ws=r.get("settings");wws=ws.split("|||");}
return wws;
}

function getWidgetDefaultSettings(id)
{
//return widget's default settings
var sets=null;
var ix=widgetMapStore.find("id",new RegExp("^"+id+"$"));
var r=widgetMapStore.getAt(ix);
if (r!=null){sets=r.get("defaults");}
return sets;
}

function removeWidget(id)
{
//remove a widget
var q=confirm(txtMsgWidgRmv);
	if (q)
	{
	var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
	var r=widgetStore.getAt(ix);
		if (r!=null)
		{widgetStore.remove(r);
		syncStore(widgetStore,"jello.widgets");
		jese.saveCurrent();}
	}
return q;
}

function editWidget(id)
{
//open the widget's settings
	customizingWidget=id;
	var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
	var r=widgetStore.getAt(ix);
		if (r==null){return;}
	
    var widget=r.get("type");
    var wht=r.get("height");
    if (wht==0 || !notEmpty(wht)){wht=150;}
    var ixx=widgetMapStore.find("id",new RegExp("^"+widget+"$"));
    var rr=widgetMapStore.getAt(ixx);
    var fn=rr.get("function");
    var ttl=rr.get("name");
    fn=fn.replace("()","("+id+",true)");
    customizationForm=eval(fn);
    customizationForm.push(new Ext.form.NumberField({fieldLabel:'Height',value:wht, width:200,invalidText:'"+txtInvalidTxt+"',id:'wgHeight'}));
    //render the form
    var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: ttl,
        height:370,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'editwidgetform',
        //iconCls:'tagformicon',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',
        items:customizationForm 
    });

    var win = new Ext.Window({
        title: txtCustWidget,
        width: 480,
        height:370,
        id:'customwidgetform',
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
            text: txtUpdate,
            tooltip:txtUpdate+' [F2]',
            listeners:{
            click: function(b,e){
            var iss=saveWidgetSettings(id);if (iss!=false){win.destroy();}}
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

function saveWidgetSettings(id)
{//save the custom settings for a widget
try{
var ix=widgetStore.find("widgetID",new RegExp("^"+customizingWidget+"$"));
var r=widgetStore.getAt(ix);}catch(e){}

if (customizingWidget==null || r==null){alert(txtInvalid);return;}

var fp=Ext.getCmp("editwidgetform");
var setString="";var setHeight=0;
var foundPipe=false;
	var getWVal=function(i,idx,lgt)
	{
  		if (i.getId()=="wgHeight")
  		{setHeight=i.getValue();}
  		else
  		{
        if (i.getId()!="folfolder")
        {  
          var vl=i.getValue();
      		try{
      		var ispipe=vl.search("\\|");
      		if (ispipe>-1){foundPipe=true;}}catch(e){}
         	setString+=vl+"|||";
        }
     	}
	};


fp.items.each(getWVal);
if (foundPipe){alert(txtAlPipeChar);return false;}
r.beginEdit();
r.set("settings",setString);
r.set("height",setHeight);
r.endEdit();
syncStore(widgetStore,"jello.widgets");
jese.saveCurrent();
customizingWidget=null;
setTimeout(function(){doWidgetRefresh(Ext.getCmp("w"+id),id);},6);	

}

function colexpWidget(p,state,id)
{
//save widget collapse/expand state
var id=p.getId();
id=id.replace("w","");
var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);
	if (r!=null)
	{
	r.beginEdit();
	r.set("collapsed",(state==0));
	r.endEdit();
	syncStore(widgetStore,"jello.widgets");
	jese.saveCurrent();
	}
		if(p.body.dom.innerText=="")
		{doWidgetRefresh(p,id);}
}

function updateWidgetOrder()
{
//update order of all widgets
status=txtMsgProc;
var pp=Ext.getCmp("portalpanel");
var cols=pp.items.length;

  for (var a=0;a<cols;a++)
  {
  //loop through columns
    var wds=pp.items.items[a].items.length;
      for (var b=0;b<wds;b++)
      {
      //loop through widgets
      var wig=pp.items.items[a].items.items[b];
      var wiID=wig.id;
      updateOrder(wiID,a,b);
      }
  }
  status="";
}

function updateOrder(id,col,row)
{
//update order of a widget
id=id.replace("w","");
var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
var r=widgetStore.getAt(ix);
	if (r!=null)
	{
	r.beginEdit();
	r.set("column",col);
	r.set("row",row);
	r.endEdit();
	syncStore(widgetStore,"jello.widgets");
	jese.saveCurrent();
	}
}

function resizePortlet(p,aw,ah,w,h)
{


		for (var x=0;x<p.items.length;x++)	
		{
		
    var pit=p.items.items[x];
    var wid=p.getId().replace("w","");
    var ht=getWidgetHeight(wid);
    
    if (ht==null){return;}
      
      setTimeout(function(){
      try{
      pit.setWidth(w-5);pit.setHeight(ht);
      }catch(e){}
      var marg=0;
      if (OLversion>=12 && jello.addSpaceOlView==true){marg=OL2007VwMargin;}
      try{var ov=document.getElementById("olv:"+wid);ov.style.height=ht-(29+marg);}catch(e){}
      
      },5);
      
    }
    
    p.doLayout();
        
    
}

function getWidgetHeight(id)
{//return saved height of widget
	var ix=widgetStore.find("widgetID",new RegExp("^"+id+"$"));
	var r=widgetStore.getAt(ix);
			if (r==null){return defaultWidgetHeight;}
		var wtype=r.get("id");
		if (wtype=="IMP" || wtype=="WFL"){return null;}
    var wht=r.get("height");
    if (wht==0 || !notEmpty(wht)){wht=defaultWidgetHeight;}
    return wht;
}

function customizeWidgetScreen()
{
//set home screen columns
var f="olv:";for (x=0;x<10;x++){try{document.getElementById(f+x).style.visibility="hidden";}catch(e){}}
var msg=txtMsgCustWidgetCols;
var buts={ok:'1 '+txtColumn,yes:'2 '+txtColumns,no:'3 '+txtColumns,cancel:txtCancel};
          Ext.Msg.show({
           title:txtHomeColumns,
           msg: msg,
           buttons: buts,
           fn: function(b,t){setWidgetColumns(b);},
           animEl: 'elId',
           icon: Ext.MessageBox.QUESTION
        });

}

function setWidgetColumns(b)
{
var f="olv:";for (x=0;x<10;x++){try{document.getElementById(f+x).style.visibility="visible";}catch(e){}}
if(b=="cancel"){return;}
var newColumns=2;
	if (b=="ok"){newColumns=1;}
	if (b=="yes"){newColumns=2;}
	if (b=="no"){newColumns=3;}
var oldColumns=jello.widgetColumns;
jello.widgetColumns=newColumns;
jese.saveCurrent();

	//loop through widgets and denominate
thisCol=0;
		var denominateWidgets=function(r)
		{	
				r.beginEdit();
				r.set("column",thisCol);
				r.endEdit();
		thisCol++;
		if(thisCol>(newColumns-1)){thisCol=0;}
		 };

widgetStore.each(denominateWidgets);
syncStore(widgetStore,"jello.widgets");
jese.saveCurrent();
Ext.info.msg(txtHomeColumns,txtMsgHomeColumns);	
pHome();
}

function resetWidgets(bup)
{   //reset widget settings and reload
if (bup){backupSettings();}
jello.widgets=[];
jese.saveCurrent();
setTimeout(function(){location.reload();},6);
}

function doWidgetRefresh(panel,pid)
{
	if (panel.collapsed && panel.title!=txtWidTImportant){return;}
	var ix=widgetStore.find("widgetID",new RegExp("^"+pid+"$"));
	var r=widgetStore.getAt(ix);
	var tp=r.get("type");
	var ht=r.get("height");
	ix=widgetMapStore.find("id",new RegExp("^"+tp+"$"));
	var r=widgetMapStore.getAt(ix);
	var fun=r.get("function");
	var nm=r.get("name");
	

	if(nm == txtTicklers)
    {var ret=eval(fun.replace("()","("+pid+",null,'')"));}
	else
	if(nm == txtWidTJTagAct)
    {//var ret=eval(fun.replace("()","("+pid+",null,'')"));
	pHome();return;
	}
    else
	{var ret=eval(fun.replace("()","("+pid+")"));}
	
    
	
	var ww=eval(ret);
	panel.removeAll(true);
	panel.add(ww);
	panel.doLayout();
	
	var pit=panel.items.items[0];
	if (ht==0){ht=150;}
	try{pit.setHeight(ht);}catch(e){}
	try{pit.setWidth(panel.getWidth()-5);}catch(e){}
	try{pit.doLayout();}catch(e){}
	
	
	setTimeout(function(){
	resizePortlet(panel);
	 },5);
 
	
}
