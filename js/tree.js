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


var tvReload={
  id:'refresh',
  handler: function(e, target, panel){
  Ext.getCmp("tvpanel").expand(true);
  try{var rn=Ext.getCmp("tree").root;
  rn.removeAll();
  }catch(e){}
  renderTree();
  }
  };

function renderTree(){
var rootNode=Ext.getCmp("tree").root;
var shc=new Ext.tree.TreeNode({text:txtShortcuts,id:'shortcuts',draggable:false,expandable:true,icon:imgPath+'action_go.gif',expanded:true,listeners:{expand:expNode}});
rootNode.appendChild(shc);

var tgs=new Ext.tree.TreeNode({text:txtJObject,id:'tags',draggable:true,expandable:true,icon:imgPath+'icon_extension.gif',expanded:true,listeners:{expand:expNode}});
rootNode.appendChild(tgs);


//query shortcuts
for (var x=0;x<shortcutStore.getCount();x++)
{
 var r=shortcutStore.getAt(x);
 var sid=r.get("id");
 var cm=r.get("cmd");
 var shn=r.get("shortcut");
 var stxt=shn+getOLFolderCounters(cm);
 var nd=new Ext.tree.TreeNode({id:"s:"+sid,text:stxt,draggable:false,expandable:false,leaf:false,icon:getShortcutIcon(r)});
 shc.appendChild(nd);
}

//filter root items
tagStore.filter("parent","0");
//add root items
for (var x=0;x<tagStore.getCount();x++)
{
 var r=tagStore.getAt(x);
 var nd=new Ext.tree.TreeNode({id:r.get("id"),text:Ext.util.Format.htmlEncode(r.get("tag")),draggable:true,expandable:true,icon:getNodeIcon(r) ,listeners:{expand:expNode}});
 var pj=r.get("isproject");
 if (pj){nd.setText("<span class=nodeprj>"+nd.text+"</span>");}
 tgs.appendChild(nd);
}
tagStore.clearFilter();


//add outlook
var tgs=new Ext.tree.TreeNode({text:txtBookmarks,id:'ol',draggable:false,expandable:true,icon:imgPath+'database.gif',expanded:false,listeners:{expand:expOLNode}});
rootNode.appendChild(tgs);
}


function expNode(ndd){
if (ndd.childNodes.length>0){return;}
var nid=ndd.id;
//filter parent items
tagStore.filter("parent",new RegExp("^"+nid+"$"));
tagStore.sort("tag","ASC");
//add node items
for (var x=0;x<tagStore.getCount();x++)
{
 var r=tagStore.getAt(x);
 var nd=new Ext.tree.TreeNode({id:r.get("id"),text:Ext.util.Format.htmlEncode(r.get("tag")),draggable:true,expandable:true,icon:getNodeIcon(r),listeners:{expand:expNode}});
 ndd.appendChild(nd);
}
tagStore.clearFilter();
}

function getNodeIcon(r)
{ 
var ic="";
	if (Ext.num(r,0)>0)
	{
	var ix=tagStore.find("id",new RegExp("^"+r+"$"));
	var rr=tagStore.getAt(ix);


	r=rr;
	}
//return the node's icon
var custicon=r.get("icon");
//custom icon overrides
if (notEmpty(custicon)){ic=imgPath+custicon;return ic;}

var tp=r.get("istag");
var tg=r.get("tag");
var tj=r.get("isproject");
var id=r.get("id");
var prv=r.get("private");
var cls=r.get("closed");
var arc=r.get("archived");
if (tp==true){ic=imgPath+"list_components.gif";}
if (tj==true){ic=imgPath+"project.gif";}
if (prv==true){ic=imgPath+"page_lock.gif";}
if (prv==true && !tp){ic=imgPath+"folder_lock.gif";}
if (tg==txtSystem){ic=imgPath+"folder_lock.gif";}
if (arc){ic=imgPath+"page_package.gif";}
if (cls){ic=imgPath+"check.gif";}
switch (id)
{
case 6:ic=imgPath+"isnext.gif";break;
case 7:ic=imgPath+"list_review.gif";break;
case 8:ic=imgPath+"list_waiting.gif";break;
case 9:ic=imgPath+"list_someday.gif";break;
}

return ic;
}

function getNodeIcon2(istag)
{
var ic="";
//return the node's icon
if (istag==true){ic=imgPath+"list_components.gif";}
return ic;
}


function nodeClick(nd,e)
{
//click of a node
var data=nodeData(nd.id);
if (data[0]=="tag")
{
  if (data[1]==0){showTagFolder(null);}
  else
  {
  var ix=tagStore.find("id",new RegExp("^"+nd.id+"$"));
  var r=tagStore.getAt(ix);
  try{
  var isTag=r.get("istag");
    if (isTag==true){lastOpenTagID=nd.id;showContext(nd.id,1);}
    else{showTagFolder(r);}
    }catch(e){}
  }
}
  if (data[0]=="sho")
  {
  //shortcuts
  if (data[1]=="0"){return;}
  var ix=shortcutStore.find("id",new RegExp("^"+data[1]+"$"));
  var r=shortcutStore.getAt(ix);
  var cmd=r.get("cmd");
try{eval(cmd);}catch(e){
    try{init();eval(cmd);}catch(f){alert(e.description);}}
  }
  if (data[0]=="olk")
  {
  if (data[1]!="0"){outlookView(data[1]);}
  }


}

//----------------outlook folders------------------------
//****outlook folder by justausr (modified for j5)
function enumerateFoldersInStores()
{    // show the various PSTs in this Outlook Session
	// skip public folders
	var colStores ;
    var oStore ;
    var oRoot ;
    var list="";
    var ret=new Array();

	colStores = NSpace.Folders;
	if(colStores.Count <= 0)
	{
		alert("colstores <=0");
		return;
	}
	// build array or name/EntryID pairs and sort on name
	var elems = new Array();
	for( var i=1; i<=colStores.Count; i++){
	  try{
		oRoot = colStores(i);
		var x = new Array(2);
		x[0] = oRoot.FolderPath.substring(2);
		x[1] = oRoot.EntryID;
		elems.push(x);
		}catch(e){Ext.info.msg(txtOLFolAccErr,txtOLFolAccErr2+' '+colStores(i).Name);}

	}
	elems.sort(tmpFolderNameSort);
	// now display in sorted order
    for( var i=0; i< elems.length; i++)
	{
		var x = new Array(2);
		x = elems[i];

		if( x[0].indexOf("\\\\Public") == 0)
			continue;
			var ff=new olFolderObject();
			ff.id=x[1];ff.name=jelloOLFName(x[0],x[1]);ff.icon="folder.gif";
			ret.push(ff);
    }
	return ret;
}



function  enumerateFolders(oFolderID)
{   //**j5

	var folders ;
    var Folder ;
    var foldercount;
	var ret=new Array();

	// if null or undefined show the root folder list
	if (oFolderID=="000"){oFolderID=null;}
	if( typeof(oFolderID) == "undefined"  || oFolderID == null || oFolderID.toLowerCase() == "mapi" || oFolderID == "")
	{
		// if only one PST, then just display children, otherwise the PST list is first
		if( NSpace.Folders.Count == 1)
		{
			oFolderID =  NSpace.GetDefaultFolder(6).Parent.EntryID;
		}
		else
			return enumerateFoldersInStores();

	}
	// get the folder from the ID
    var oFolder = NSpace.GetFolderFromID(oFolderID);
	// was a parent supplied?  No get one
	  // are there any children
    foldercount = oFolder.folders.Count;
    // loop through them build array of names/index and sort
    if( foldercount > 0){
		folders = oFolder.folders;
		var parstrng ="";
		if( typeof(oFolderID) != "undefined")
			parstring = ",\"" + oFolderID + "\"";
		var elems = new Array();
		for( var i=1; i <= foldercount; i++){
			var x = new Array(2);
			var y = folders(i).FolderPath.lastIndexOf("\\");
			x[0] = folders(i).FolderPath.substring(y+1);
			x[1] = i;
			elems.push(x);
		}

/* search folder routine. Still needs work
  if(oFolder.Parent=="Mapi" && OLversion>=12)
  {//for outlook 2007
  var sf=ol.Session.Stores.Item(oFolder.Name).GetSearchFolders;
  var sfc=sf.Count;
  		for( var i=1; i <= sfc; i++){
			var x = new Array(3);
			var y = sf.Item(i).FolderPath.lastIndexOf("\\");
			x[0] = sf.Item(i).FolderPath.substring(y+1);
			x[1] = i;
			x[2] =true;
			elems.push(x);
            }

  }
  */
		elems.sort(tmpFolderNameSort);
		// now do the work, we use the name sorting order and get the index of the element to use from the name sort array
        for( var i=1; i <= foldercount; i++){

			var x = elems[i-1];
			var thisfold = folders(x[1]);
			var fname = thisfold.FolderPath.lastIndexOf("\\");
			var typ = thisfold.DefaultItemType;
			var thisimg="";
			if( typ == 1)
				thisimg = icApp;
			if( typ == 2  || typ == 7)
				thisimg = icContact;
			if( typ == 3)
				thisimg = icTask;
			if( typ == 4)
				thisimg = icJournal;
			if( typ == 5)
				thisimg = icNote;
			if( typ == 0)
				thisimg = icMail;
	//		if (x.length>2){thisimg="sfolder.gif";}
			var ff=new olFolderObject();
			ff.id=thisfold.EntryID;ff.name=jelloOLFName(thisfold,thisfold.EntryID);ff.icon=thisimg;
			ret.push(ff);
		}

	}
	return ret;
}
function tmpFolderNameSort(a,b)
{
	var x = a[0].toLowerCase();
	var y = b[0].toLowerCase();

	return ((x < y) ? -1 : ((x > y) ? 1 : 0));
}

function olFolderObject()
{
this.name="";this.id="";this.icon="";
}

function expOLNode(nd)
{
if (nd.childNodes.length>0){return;}
var o=nd.id.split(":");
var id=o[1];
var children=enumerateFolders(id);
	for (var x=0;x<children.length;x++)
	{
	var of=children[x];
	var od=new Ext.tree.TreeNode({id:"o:"+of.id,text:of.name,draggable:true,expandable:true,icon:'img/'+of.icon,listeners:{expand:expOLNode}});
	nd.appendChild(od);
	}
}

function outlookView(fid)
{
//view of an outlook folder

var ix=olFolderStore.find("eid",new RegExp("^"+fid+"$"));
r=olFolderStore.getAt(ix);
if (r!=null)
{
var nw=r.get("newwindow");
if (nw==true){OLFolderOpenNewWindow(fid);return;}
}

var f=NSpace.GetFolderFromID(fid);
pInbox(f);

}


function addTreeTag(parent)
{
//add a tag node
var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();

if (nd==null){editTag(null,false,0);return;}

    if (parent>0){var ix=tagStore.find("id",new RegExp("^"+parent+"$"));var r=tagStore.getAt(ix);
    var isTag=r.get("istag");
    editTag(null,isTag,parent);}
else{editTag(null,false,parent);}
}

function popUpClick(bt,nd)
{
var id=nd.id;
var data=nodeData(id);

  if (bt=="edit-node")
    {
    if(data[0]=="tag"){editTag(data[1]);}
    if(data[0]=="olk"){editOLFolder(data[1]);}
    }

  if (bt=="insert-node")
    {
    if(data[0]=="tag"){addTreeTag(data[1]);}
    if(data[0]=="olk"){olAddFolder(data[1],nd);}
    }

  if (bt=="ol-open")
    {
    if(data[0]=="olk"){OLFolderOpenNewWindow(data[1]);}
    }

  if (bt=="tg-icon")
    {
    if(data[0]=="tag"){customIconTagForm(data[1]);}
    }

  if (bt=="add-shcut")
  {
  //lastToShortcutParent=Ext.getCmp("tree").getNodeById("shortcuts");
  newShortCut(data[0],data[1],nd.text,nd,true);
  }

  if (bt=="delete-node")
    {
    if(data[0]=="tag"){if (choice=confirm(txtMsgDelItem)==false){return;}deleteTag(data[1],nd);}
    if(data[0]=="sho"){if (choice=confirm(txtMsgDelItem)==false){return;}deleteShortcut(data[1],nd);}
    if(data[0]=="olk"){olDelFolder(data[1],nd);}
    }
   if( bt == "tag-node" && data[0] == "tag")
   {
	  var ttg=tagStore.getAt(tagStore.find("id",new RegExp("^"+data[1]+"$")));

	  tasksWithTag(null, ttg.get("tag"));
   }
}

function treeMenu(node,e)
{
            node.select();
            var data=nodeData(node.id);
            var c = node.getOwnerTree().contextMenu;
            c.contextNode = node;
            if (data[0]=="sho"){Ext.getCmp("tg-icon").hide();Ext.getCmp("ol-open").hide();Ext.getCmp("add-shcut").disable();Ext.getCmp("insert-node").disable();Ext.getCmp("edit-node").disable();Ext.getCmp("delete-node").enable();Ext.getCmp("tag-node").hide();}
            if (data[0]=="tag"){Ext.getCmp("tg-icon").show();Ext.getCmp("ol-open").hide();Ext.getCmp("add-shcut").enable();Ext.getCmp("insert-node").enable();Ext.getCmp("edit-node").enable();Ext.getCmp("delete-node").enable();Ext.getCmp("tag-node").enable();}
            if (data[0]=="olk"){Ext.getCmp("tg-icon").hide();Ext.getCmp("ol-open").show();Ext.getCmp("add-shcut").enable();Ext.getCmp("insert-node").enable();Ext.getCmp("edit-node").enable();Ext.getCmp("delete-node").enable();Ext.getCmp("tag-node").hide();}
            if (data[1]=="0"){Ext.getCmp("tg-icon").hide();Ext.getCmp("ol-open").hide();Ext.getCmp("insert-node").disable();Ext.getCmp("edit-node").disable();Ext.getCmp("delete-node").disable();Ext.getCmp("add-shcut").disable();Ext.getCmp("tag-node").hide();}
            if (data[0]=="tag"){Ext.getCmp("insert-node").enable();}
            c.showAt(e.getXY());
}

function nodeData(ndid)
{
var dt=new Array();
var ntype="";var nid=0;
if(Ext.num(ndid,-1)>0){ntype="tag";nid=ndid;}
if (ndid=="shortcuts"){ntype="sho";nid=0;}
if (ndid=="tags"){ntype="tag";nid=0;}
if (ndid=="ol"){ntype="olk";nid=0;}

  if (typeof(ndid)!="number")
  {
  var oth=ndid.search(":");
  if (oth>-1){
  var o=ndid.split(":");
  if (o[0]=="s"){ntype="sho";nid=o[1];}
  if (o[0]=="o"){ntype="olk";nid=o[1];}
  }
  }

 dt.push(ntype);
 dt.push(nid);
 return dt;
}

function nodeDrag(obj)
{
//check for drag and drop
var data=nodeData(obj.data.node.id);
var tdata=nodeData(obj.target.id);

if (obj.data.node.contains(obj.target)){obj.cancel=true;}
if (data[0]=="tag" && tdata[0]=="olk"){obj.cancel=true;}
if (data[0]=="olk" && tdata[0]=="tag"){obj.cancel=true;}
if (data[0]=="olk" && tdata[0]=="olk"){obj.cancel=true;}
if (data[0]=="tag" && tdata[0]=="tag")
{if (getTagType(data[1])==false && getTagType(tdata[1])==true){obj.cancel=true;}}
if (data[0]=="tag")
//{if (getTagSys(data[1])==true){obj.cancel=true;}}
if (data[0]=="tag" && tdata[0]=="sho"){lastToShortcutParent=obj.data.node.parentNode;}
if (data[0]=="olk" && tdata[0]=="sho"){lastToShortcutParent=obj.data.node.parentNode;}
}

function nodeDrop(obj)
{
//node dropped!
var data=nodeData(obj.data.node.id);
var tdata=nodeData(obj.target.id);
  if (data[0]=="tag" && tdata[0]=="tag")
  {
		if (tdata[1]!=null || tdata[1]!="")
		{
		//tag move
		var ix=tagStore.find("id",new RegExp("^"+data[1]+"$"));
		var r=tagStore.getAt(ix);
		r.beginEdit();
		r.set("parent",tdata[1]);
		r.endEdit();
    tagStore.clearFilter();
		syncStore(tagStore,"jello.tags");
		jese.saveCurrent();
		}
		else
		{
		Ext.info.msg(txtError,txtMsgNoMoveErr);
		obj.cancel=true;
		}
  }

if (tdata[0]=="sho")
{
  newShortCut(data[0],data[1],obj.data.node.text,obj.data.node);
}

}

function newShortCut(key,id,text,toNode,noDD)
{
  if (key=="tag")
  {
  //new shortcut of a tag
  var thisimg="list_components.gif";
  var ttp=getTagType(id);
  var tpri=getTagPrivacy(id);
  var tarc=getTagArchived(id);
  var tic=getTagIcon(id);
  var tpj=getTagProject(id);
  if (!ttp){thisimg="ffolder.gif";}
  if (tpj){thisimg="project.gif";}
  if (tpri && ttp){thisimg="page_lock.gif";}
  if (tpri && !ttp){thisimg="folder_lock.gif";}
  if (tarc){thisimg="page_package.gif";}
  if (notEmpty(tic)){thisimg=tic;}
  var added=newShortcut("showContext("+id+",1)",text,thisimg);
  }

  if (key=="olk")
  {
  //new shortcut of an outlook folder
  var thisfold=NSpace.GetFolderFromID(id);
  			var typ = thisfold.DefaultItemType;
			var thisimg="folder.gif";
			if( typ == 1)
				thisimg = icApp;
			if( typ == 2  || typ == 7)
				thisimg = icContact;
			if( typ == 3)
				thisimg = icTask;
			if( typ == 4)
				thisimg = icJournal;
			if( typ == 5)
				thisimg = icNote;
			if( typ == 0)
				thisimg = icMail;

  var added=newShortcut("outlookView('"+id+"')",text,thisimg);
  }
  if (noDD==null){

  try{
  lastToShortcutParent.appendChild(toNode);
  lastToShortcutParent.expand();
      }catch(e){}
  }
}