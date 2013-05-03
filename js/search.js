// search Module
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


function finderFocus()
{
finder.value="";
var f=Ext.get("finder");
f.highlight();
}

function finderBlur()
{
finder.value=txtFind+"...";
}

function quickSearch()
{
//perform a quick search on actions and tags/folders
var ss=finder.value;

if (!notEmpty(ss) || ss==txtFind+"...(ctrl+Z)" || ss==txtFind+"..."){return;}
var ctx=txtSearchRes+" (" + ss +")";
status=txtMsgStSearching+"...";
//search actions
dasl="urn:schemas:httpmail:subject LIKE '%"+ss+"%' OR urn:schemas:httpmail:subject LIKE '"+ss+"%' OR urn:schemas:httpmail:subject LIKE '%"+ss+"' OR urn:schemas:httpmail:subject LIKE '@%"+ss+"%'";
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var its=iF.Restrict("@SQL="+dasl);
counter=its.Count;


//search tags
var oItems=new Array();
tagStore.filter("tag",ss);
for (var x=0;x<tagStore.getCount();x++)
  {
  var r=tagStore.getAt(x);
  oItems.push(r);
  }
  tagStore.clearFilter(); 

//search journal

try{var jif=NSpace.GetFolderFromID(jello.nonActionFolder).Items;
var jits=jif.Restrict("@SQL="+dasl);
jcounter=jits.Count;
  for (var x=1;x<=jcounter;x++)
  {
  var jit=jits(x);
  var ar=new actionObject(jit,true);
  var newRec=new actionRecord(ar);
  oItems.push(newRec);
  }
 }catch(e){}
 
 
buildGrid(ctx,its,counter,null,oItems,true);
lastOpenTagID=0;lastContext=null;
Ext.getCmp("inlinenewaction").hide();
Ext.getCmp("filter").hide();
status="Ready";
}

function calcLatest()
{
//create or show the latest tasks link
var t=Ext.get("topLatest");
t.highlight("#ff0000", {
    attr: "background-color",
    endColor: "ffffff",
    easing: 'easeIn',
    duration: 2
});	
}

function pLatest()
{
//popup list with latest actions created
showOLViewCtl(false);
var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: txtTtlLatestAt,
        height:370,
        bodyStyle:'padding:5px 5px 0 5px',
        autoScroll:true,
        floating:false,
        labelWidth:480,
        id:'latestform',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',
       	html:pLatestList()
       

    });

    var win = new Ext.Window({
        title: txtLatest,
        width: 480,
        height:370,
        id:'thelatestform',
        minWidth: 300,
        minHeight: 200,
        autoScroll:true,
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
    
    var map = new Ext.KeyMap('thelatestform', {
    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
            try{win.destroy();}catch(e){}
            
            },
    scope: this});

}

function pLatestList(retLast)
{
var ret="<table width=100% cellpadding=3 cellspacing=0>";
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var its=iF.Restrict("@SQL=(%last7days(urn:schemas:calendar:created)%)");
its.Sort("urn:schemas:calendar:created",true);
  if(its.Count>0){
  if (retLast==true){return its(1);}
  var lmt=10;
  if (its.Count<10){lmt=its.Count;}
    for (var x=1;x<=lmt;x++)
    {
 //   var tList=tagList(null,its(x).itemProperties.item(catProperty).Value);
    ret+="<tr><td class=printChars width=1%><img src=img//task.gif></td><td width=1% class=printChars>"+DisplayDate(new Date(its(x).CreationTime))+"</span></td><td width=98% class=printChars style='padding-left:10px;'><a class=jellolink onclick=scAction('"+its(x).EntryID+"')>"+its(x).Subject+"</a></td></tr>";
    }  
  }
  if (retLast==true){return null;}
  ret+="</table>";
return ret;
}

function theLatestThing()
{
//return the latest action for showing on the sidebar
var ret="";
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
var its=iF.Restrict("@SQL=(%last7days(urn:schemas:calendar:created)%)");
its.Sort("urn:schemas:calendar:created",true);
  if(its.Count>0){
    var tlt=its(1).Subject;
    if (tlt.length>20){tlt=tlt.substr(0,20)+"...";}
    ret+="<table width=100%><tr><td><img src=img//task.gif></td><td width=98% padding-left:3px;'><a title='Ctrl+Alt+L / "+txtClickEdit+"' class=jellolink style='font-size:70%;'onclick=scAction('"+its(1).EntryID+"')>"+tlt+"</a></td></tr></table>";
  }
return ret;
}

function updateTheLatestThing()
{
var lt=document.getElementById("thelatest");
var ltt=theLatestThing();
lt.innerHTML=ltt;
}

