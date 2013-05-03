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


var lastToShortcutParent="";

var shortcutStore=new Ext.data.JsonStore({
        fields: [
           {name: 'id'},
           {name: 'shortcut'},
           {name: 'cmd'},
           {name: 'sys'},
           {name: 'icon'}
        ],
        data:jello.shortcuts
    });

var shortcutRecord = Ext.data.Record.create([
    {name: 'id',type: 'int'},
	{name: 'shortcut'},
	{name: 'cmd'},
	{name: 'sys',type: 'boolean'},
	{name: 'icon'}
]);

function getShortcutIcon(r)
{
var ic=imgPath+"icon_link.gif";
//return the node's icon
var tp=r.get("icon");
if (notEmpty(tp) || tp!="0"){ic=imgPath+tp;}
return ic;
}

function shortcutExists(cmd)
{
//check if a shortcut command exists
shortcutStore.filter("cmd",cmd);
var ret=false;
if (shortcutStore.getCount()>0){ret=true;}
shortcutStore.clearFilter(); 
return ret;
}

function newShortcut(cmd,ndname,img)
{

  jello.lastShortcutId++;
  var cc=cmd;
  var exists=false;var added=false;

if (shortcutExists(cc)){exists=true;}

if (exists==false)
{//create shortcut
if (!notEmpty(img)){img='icon_link.gif';}
if (cmd.substr(0,11)=="outlookView")
{try{var x=ndname.search("</span>");if (x>-1){ndname=ndname.substr(0,(x+7));}}catch(e){}}
var sr=new shortcutRecord({
id:jello.lastShortcutId,
shortcut:''+ndname,
cmd:cc,
sys:false,
icon:img
});
shortcutStore.add(sr);
syncStore(shortcutStore,"jello.shortcuts");
jese.saveCurrent();
added=true;
}

//update tree
if (Ext.getCmp("tvpanel").body.dom.innerText!='')
{
if (added==true ){
//add new shortcut node
var shc=Ext.getCmp("tree").getNodeById("shortcuts");
var nd=new Ext.tree.TreeNode({id:"s:"+jello.lastShortcutId,text:ndname,draggable:false,expandable:false,leaf:false,icon:getShortcutIcon(sr)});
shc.appendChild(nd);}else{alert(txtPromptShortcutExists);}
}
return added;

}

function deleteShortcut(id,node)
{
//delete a shortcut, remove node
var ix=shortcutStore.find("id",new RegExp("^"+id+"$"));
var r=shortcutStore.getAt(ix);

var shName=r.get("shortcut");


//find the shortcut and delete it
shortcutStore.remove(r);
  
//save changes
  syncStore(shortcutStore,"jello.shortcuts");
  jese.saveCurrent();    

  if (node!=null)
  {//if delete was called from a tree...
  node.remove();
  gridRemoveTag("s:"+id);
  } 
  
}
function alShortcut(id,itype)
{
//create shortcut from action list

if (itype=="t")
{
var gtext=getTagName(id);
var cmd="";var img=null;
  var ix=tagStore.find("id",new RegExp("^"+id+"$"));
  var r=tagStore.getAt(ix);
  try{var isTag=r.get("istag");
      var prv=r.get("private");
      var arc=r.get("archived");
    if (isTag){cmd="showContext("+id+",1)";}
    if (prv){img="page_lock.gif";}
    if (arc){img="page_package.gif";}
    else{cmd="showTagFolder(null,"+id+")";}}catch(e){}
//this does not show tag folders
var added=newShortcut(cmd,gtext,img);
}

if (itype=="a")
{
var gtext=NSpace.GetItemFromID(id).Subject;
var added=newShortcut("scAction('"+id+"')",gtext);
}
}

function scAction(id)
{
//open action from shortcut
try{Ext.getCmp("thelatestform").destroy();}catch(e){}
var it=NSpace.GetItemFromID(id);
var ar=new actionObject(it);
var newRec=new actionRecord(ar);
var ai=new Array();
ai.push(newRec);
editAction(ai,true);
}