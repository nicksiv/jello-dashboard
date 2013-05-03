//------------------- tag -------------
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


//-------------------------------------
var tagStore=new Ext.data.JsonStore({
        fields: [
           {name: 'id'},
           {name: 'parent'},
           {name: 'tag', type:'string'},
           {name: 'istag'},
           {name: 'notes'},
           {name: 'archive'},
           {name: 'color'},
           {name: 'icon'},
           {name: 'isdoc'},
           {name: 'private'},
           {name: 'archived'},
           {name: 'sys'},
           {name: 'isproject'},
           {name: 'path'},
           {name: 'filter'},
           {name: 'closed'},
           {name: 'modified', type:'date'}
          
        ],
        data:jello.tags
        
    });


//define tagRecord
 var tagRecord = Ext.data.Record.create([
    {name: 'id',type: 'int'},
    {name: 'parent',type: 'int'},
	{name: 'tag'},
	{name: 'istag',type: 'boolean'},
	{name: 'notes'},
 {name: 'archive'},
 {name: 'color'},
 {name: 'icon'},
 {name: 'isdoc',type:'boolean'},
	{name: 'private',type: 'boolean'},
	{name: 'archived',type: 'boolean'},
	{name: 'sys',type: 'boolean'},
	{name: 'isproject',type: 'boolean'},
	{name: 'path'},
	{name: 'filter'},
	{name: 'closed'},
  {name: 'modified', type:'date'}

]);

 var tagComboR = Ext.data.Record.create([
	{name: 'tag'}
]);


//Outlook folder definitions store & record
var olFolderStore=new Ext.data.JsonStore({
        fields: [
           {name: 'eid'},
           {name: 'fname'},
           {name: 'activex'},
           {name: 'newwindow'},
           {name: 'counter'},
           {name: 'defolview'}
        ],
        data:jello.outlookFolders
    });

var olFolderRecord = Ext.data.Record.create([
    {name: 'eid'},
	{name: 'fname'},
	{name: 'activex',type:'boolean'},
	{name: 'newwindow',type:'boolean'},
	{name: 'counter',type:'boolean'},
	{name: 'defolview'}
]);

//tag manager filters
var tFilters=[
        ['0',txtAll],
        ['5',txtTags],
        ['1',txtProjects],
        ['2',txtFolders],        
        ['3',txtInactive],
        ['4',txtMarkContextPrivate]
        
        ];        

var tFilterValues=new Ext.data.SimpleStore({
        fields: [
           {name: 'value'},
           {name: 'text'}
        ],
        data:tFilters
    });

function getTagsArray(){
//j5
//return an array of all tags
tagStore.clearFilter();
tagStore.filter("istag",true);
var arr=new Array();
	for (var x=0;x<tagStore.getCount();x++)
	{
		var tn=tagStore.getAt(x).get("tag");
    //var isArc=tagStore.getAt(x).get("archived");
    var isClosed=tagStore.getAt(x).get("closed");
    if (notEmpty(tn) && isClosed!=true){arr.push(Ext.util.Format.htmlEncode(tn));}
	}
globalTags=arr;
tagStore.clearFilter();
}


function tagDescription(ctx)
{
try{var r=getTag(ctx);
var fnotes=r.get("notes");
var ctxid=r.get("id");
var ret="";
if (notEmpty(fnotes)){ret+="&nbsp;<span style=font-size:9px;color:black;><em>("+fnotes.substr(0,100)+")</em></span>";}
ret+="&nbsp;<a title='"+txtClickEdit+" (Ctrl+T)' class=jellolink onclick=editTag("+ctxid+"); style='font-weight:normal;text-decoration: underline;'>["+txtEdit.toLowerCase()+"]</a>";
return ret;}catch(e){return "";}
}

function getTag(tg){
tg=Ext.util.Format.htmlDecode(tg);
tg=Ext.escapeRe(tg);
try{var ix=tagStore.find("tag",new RegExp("^"+tg+"$","i"));
var r=tagStore.getAt(ix);
return r;}catch(e){return null;}
}

function getTagByID(id){
var ix=tagStore.find("id",new RegExp("^"+id+"$"));
var r=tagStore.getAt(ix);
try{
return r;}catch(e){return null;}
}

function getTagArcFolder(tg){
tg=Ext.escapeRe(tg);
var ix=tagStore.find("tag",new RegExp("^"+tg+"$","i"));
var r=tagStore.getAt(ix);
var rt="";
try{rt=r.get("archive");}catch(e){}
return rt;
}

function getTagName(tid){
if (tid==0){return txtJObject;}
if (tid==-1){return "##"+txtInbox;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
	if (r!=null)
	{return r.get("tag");}
	else
	{return null;}
}

function getTagID(tn){
var tid=tn.replace(new RegExp("~","g")," ");
//var tid=tid.replace("+","\+");
//var tid=tid.replace("[","\[");
tid=Ext.escapeRe(tid);
var ix=tagStore.find("tag",new RegExp(tid,"i"));
var r=tagStore.getAt(ix);

if (r!=null){return r.get("id");}
else{return 0;}
}

function getTagPrivacy(tid){
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
try{return r.get("private");}catch(e){return null;}
}

function getTagArchived(tid){
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
try{return r.get("archived");}catch(e){return false;}
}

function getTagIcon(tid){
if (tid==0){return null;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
try{return r.get("icon");}catch(e){return null;}
}

function getTagProject(tid){
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
try{return r.get("isproject");}catch(e){return false;}
}


function setTagArchived(tid){
//toggle inactivity status of tag
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
  try{
  var cstat=r.get("archived");
  var cpj=r.get("isproject");
  if (!cpj){alert(txtInvalid);return false;}
  r.beginEdit();
  r.set("archived",!cstat);
  r.endEdit();
  tagStore.clearFilter();
  syncStore(tagStore,"jello.tags");
  jese.saveCurrent();
  return true;
  }catch(e){return false;}
}

function setTagPrivacy(tid){
//toggle inactivity status of tag
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
  try{
  var ctg=r.get("istag");
  var cstat=r.get("private");
  if (!ctg){alert(txtInvalid);return false;}
  r.beginEdit();
  r.set("private",!cstat);
  r.endEdit();
  tagStore.clearFilter();
  syncStore(tagStore,"jello.tags");
  jese.saveCurrent();
  return true;
  }catch(e){return false;}
}

function getTagType(tid){
if (tid==0){return false;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
return r.get("istag");
}

function getTagSys(tid){
if (tid==0){return true;}
var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
return r.get("sys");
}
function selectTagOnTree(r)
{//select specific tag record on hierarchy
var tree=Ext.getCmp("tree");
var nd=tree.getSelectionModel().getSelectedNode();
}

function isidtag(firstID)
{  var ret=false;
 try{if(firstID.substr(0,1)=="t"){ret=true;}}catch(e){}
 return ret;
}

function editTag(tagsitems,folderNotAllowed,toParent){
//j5
//open tag record in edit form
try{Ext.getCmp("thetagform").destroy();}catch(e){}

if (typeof(tagsitems)=="number")
{
  if (tagsitems==0){return;}
  var ttg=tagStore.getAt(tagStore.find("id",new RegExp("^"+tagsitems+"$")));
	var ar=new actionObject(ttg,true);
	var newRec=new actionRecord(ar);
	var tagsitems=new Array();
	tagsitems.push(newRec);
}

if (tagsitems!=null)
{
var id=tagsitems[0].get("entryID");
var tp=tagsitems[0].get("type");

var r=tagStore.getAt(tagStore.find("id",new RegExp("^"+id+"$")));
var fid=r.get("id");
var fparent=r.get("parent");
var vftag=r.get("tag");
var fnotes=r.get("notes");
var fistag=r.get("istag");
var fprive=r.get("private");
var farchiveID=r.get("archive");
var farchived=r.get("archived");
var fisproject=r.get("isproject");
var fisclosed=r.get("closed");
if (notEmpty(farchiveID)){try{var farchive=NSpace.GetFolderFromID(farchiveID).FolderPath;}catch(e){farchive=txtUndefined;}}
else{try{var farchive=NSpace.GetFolderFromID(jello.archiveFolder).FolderPath;}catch(e){var farchive=txtUndefined;}}
var fsys=r.get("sys");
var finfo="";

if (fsys==true){finfo+="<span class=tagnext>"+txtSysTagIfo+"</span>";}
if (farchived==true){finfo+="<span class=tagnext>"+txtArchivedInfo+"</span>";}
if (farchived==true){finfo+="<span class=tagprj>"+txtProject+"</span>";}
if (fisclosed==true){finfo+="<span class=tagprj><b> "+txtCompletedInfo+" </b></span>";}
var fttl=txtEditTag;
var tchide=true;
}
else
{
var fid=0;
var fparent=toParent;
if (toParent==null){fparent=lastOpenTagID;}
var vftag=txtNewTag;
var fnotes="";
var fistag=true;
var fprive=false;
var farchived=false;
var fsys=false;
var fisproject=false;
var fisclosed=false;
var farchiveID=jello.archiveFolder;
try{var farchive=NSpace.GetFolderFromID(farchiveID).FolderPath;}catch(e){var farchive=txtUndefined;farchiveID=null;};
var fttl=txtNewTag;
var finfo="";
var finfo=+"<span class=tagprj>"+txtNewRecord+"</span>";
var tchide=false;
} 
try{var fparName=tagStore.getAt(tagStore.find("id",new RegExp("^"+fparent+"$"))).get("tag");}catch(e){var fparName="";}
var fttype=txtTag;
if (fistag==false){fttype=txtFolder;}
if (fisproject==true){fttype=txtProject;}

 var simple = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: fttl,
        height:400,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:80,
        id:'tagform',
        iconCls:'tagformicon',
        buttonAlign:'center',
        defaults: {width: 280},
        defaultType: 'textfield',

        items: [{
                fieldLabel: txtTagName,
                id: 'ftag',
                value:vftag,
                disabled:false,
                tabIndex:0,
                allowBlank:false
            },
				new Ext.form.ComboBox({
                fieldLabel: txtType,
                id: 'ttype',
                store:new Array(txtTag,txtProject,txtFolder),
                hideTrigger:false,
                typeAhead:true,
                editable:true,
                forceSelection:true,
                listeners:{beforeselect:function(cb,r,idx){if(idx==2 && fid>0){alert(txtAlrtCantChgFld);return false;}}},
                triggerAction:'all',
                emptyText:fttype,
                disabled:false,
                mode:'local'
            }),                
				 new Ext.form.Hidden({
                id: 'fid',
                value:fid 
                }),
                new Ext.form.Hidden({
                id: 'fparent',
                value:fparent 
                }),
                new Ext.form.TextArea({
                fieldLabel: txtNotes,
                id: 'fnotes',
                height:100,
                value:fnotes
            }),
            {
                fieldLabel: txtParent,
                id: 'fpar',
                value:fparName,
                disabled:true
            },
            new Ext.form.TriggerField({
            	fieldLabel:txtJArcFol,
            	value:farchive ,
            	width:200,
            	hidden:!fistag,
            	onTriggerClick:function(e){selectTagFolder();Ext.getCmp("arcflink").update('');},
            	hideLabel:!fistag,
            	id:'farchive'
            	}),
            	new Ext.form.Label({
            	id:'arcflink',
              html:'<a class=jellolinktop onclick=OLFolderOpenNewWindow(&quot;'+farchiveID+'&quot;);>[Open Archive folder]</a>',
              cls:'formatflist'
              }),
              new Ext.form.Hidden({
            	value:farchiveID,
            	id:'farchiveID'
              }),
             new Ext.form.Checkbox({
                fieldLabel: txtMarkContextPrivate,
                id: 'fprivate',
				        checked:fprive
            }),
            new Ext.form.Checkbox({
                fieldLabel: txtInactive,
                id: 'fisarchived',
				        checked:farchived
            }),
            new Ext.form.Checkbox({
                fieldLabel: txtCompleted,
                id: 'fisclosed',
				        checked:fisclosed
            }),
            new Ext.form.Label({
				html:finfo
            })
        ]

    });



   var win = new Ext.Window({
        title: txtEdit,
        width: 480,
        height:460,
        id:'thetagform',
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
            saveTagForm(b,win);
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
  setTimeout(function(){Ext.getCmp("ftag").focus(true,100);},1);

//shortcut keys
var map = new Ext.KeyMap('thetagform', {
    key: Ext.EventObject.F12,
    fn: function(){
    win.destroy();
    },
    scope: this});   

var map2 = new Ext.KeyMap('thetagform', {
    key: Ext.EventObject.F2,
    fn: function(){
    try{
    saveTagForm(null,win);}catch(e){}},
    scope: this});   

}

function saveTagForm(b,win)
{
//save tag form

try{var id=Ext.getCmp("fid").getValue();}catch(e){id=0;}
var ttag=Ext.getCmp("ttype").getValue();

if (id>0){
//edit existing
var newValue=Ext.getCmp("ftag").getValue();
if (isValidTag(newValue)==false){return;}
var ix=tagStore.find("id",new RegExp("^"+id+"$"));
var r=tagStore.getAt(ix);
var oldValue=r.get("tag");
var tid=r.get("id");
var tparent=r.get("parent"); 
var tistag=r.get("istag");
var tisproject=r.get("isproject");
var tisclosed=Ext.getCmp("fisclosed").getValue();
var tnotes=Ext.getCmp("fnotes").getValue();
var tprive=Ext.getCmp("fprivate").getValue();
var tarchived=Ext.getCmp("fisarchived").getValue();
var tarchive=Ext.getCmp("farchiveID").getValue();
if (ttag==txtTag){tistag=true;tisproject=false;}
if (ttag==txtProject){tisproject=true;tistag=true;}
if (ttag==txtFolder){tisproject=false;tistag=false;}

r.beginEdit();
r.set("id",tid);
r.set("tag",newValue);
r.set("parent",tparent);
r.set("istag",tistag);
r.set("isproject",tisproject);
r.set("notes",tnotes);
r.set("private",tprive);
r.set("archived",tarchived);
r.set("archive",tarchive);
r.set("closed",tisclosed);
r.set("modified",new Date());

r.endEdit();
var rr=r;
tagStore.clearFilter();
syncStore(tagStore,"jello.tags");
//save settings
jese.saveCurrent();
//replace tag name in other places
updateOLCategory(oldValue,newValue);
  if (trimLow(newValue)!=trimLow(oldValue))
  {
  replaceItemTag(oldValue,newValue);
  try{
  var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();
	if (nd.text==oldValue)
  {nd.setText(newValue);}
  else{
  if (nd.firstChild!=null)
	{var nnd=nd.findChild("text",oldValue);if(nnd!=null){nnd.setText(newValue);}}
	}
    }catch(e){}
  } 
//if there is a grid update store
    try{
    var gstore=Ext.getCmp("grid").getStore();
    var ix=gstore.findBy(function(r){if(r.get("type")=="t" && r.get("entryID")==tid){return true;}});
    var r=gstore.getAt(ix);
    r.beginEdit();
    r.set("subject",newValue);
    r.endEdit();
    }catch(e){}
    if (lastOpenTagID==tid){showTagFolder(rr);}
    
}
else
{
/////// new tag	 /////////////
var newValue=Ext.getCmp("ftag").getValue();
if (isValidTag(newValue)==false){return;}
if (newValue==txtNewTag){Ext.getCmp("ftag").focus(true);return;}
var tparent=Ext.getCmp("fparent").getValue(); 
if (notEmpty(ttag)==false){ttag=txtTag;}
if (getTagType(tparent)==true && ttag==txtFolder){alert(txtMsgFolUndTag);return;}
var tistag=true;
if (ttag==txtFolder){tistag=false;}
if (ttag==txtProject){tisproject=true;}
var tnotes=Ext.getCmp("fnotes").getValue();
var tprive=Ext.getCmp("fprivate").getValue();
var tisclosed=Ext.getCmp("fisclosed").getValue();
var tarchive=Ext.getCmp("farchiveID").getValue();
var tarchived=Ext.getCmp("fisarchived").getValue();

jello.lastTagId++;
var tr=new tagRecord({
id:jello.lastTagId,
tag:newValue,
parent:tparent, 
istag:tistag,
isproject:tisproject,
notes:tnotes,
archive:tarchive,
archived:tarchived,
closed:tisclosed,
modified:new Date(),
isprivate:tprive
});
tagStore.add(tr);
tagStore.clearFilter();
syncStore(tagStore,"jello.tags");
//save settings
jese.saveCurrent();
updateOLCategory(newValue);

try{
  if (Ext.getCmp("tvpanel").body.dom.innerText!='')
  {
  var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();
  	if (nd==null)
  	{
  	//select jello objects
  	var nd=Ext.getCmp("tree").getNodeById("tags");
  	}
  	else
  	{
  	//if another type node was selected reset nd to tags
  	var data=nodeData(nd.id);
  	if (data[0]!="tag"){var nd=Ext.getCmp("tree").getNodeById("tags");}
  	}
  	if (nd.firstChild!=null || nd.expanded==true)
  	{
  	var nnd=new Ext.tree.TreeNode({text:newValue,id:jello.lastTagId,expandable:true,icon:getNodeIcon(jello.lastTagId),expanded:false,listeners:{expand:expNode}});
  	nd.appendChild(nnd);
  	}
  }
}catch(e){}

// if there is a grid, update its store	
try{
if (lastOpenTagID==tparent)
{
var store=Ext.getCmp("grid").getStore();
var ar=new actionObject(tr,true);
var newRec=new actionRecord(ar);
store.add(newRec);
}
}catch(e){}

}
Ext.info.msg(txtEditTag, txtMsgTagUpd);
try{win.destroy();}catch(e){}
}

function replaceItemTag(oldVal,newVal)
{
//replace occurences of a context or project in items collection with another
		iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
		
    its=iF.Restrict("@SQL=urn:schemas-microsoft-com:office:office#Keywords like '%"+Trim(oldVal)+"%'");

    var itsCounter=its.Count;
    	if (itsCounter>0)
			{
				for (var x=itsCounter;x>0;x--)
				{
				try{
        var e=its(x).itemProperties.item(catProperty).Value;
				e=e.replace(new RegExp(",","g"),";");
				var t=e.split(";");
				var nca="";
					for (var y=0;y<=t.length;y++)
					{
						if (notEmpty(t[y]))
						{
						   if (trimLow(t[y])==trimLow(oldVal))
						   {nca=nca+newVal+";";}
						   else
						   {nca=nca+t[y]+";";}
						}
					}
				its(x).itemProperties.item(catProperty).Value=nca;	
				its(x).Save();
				}catch(e){}
        }
			}
			
//replace journal entries tags			
		iF=NSpace.GetFolderFromID(jello.nonActionFolder).Items;
		
    its=iF.Restrict("@SQL=urn:schemas-microsoft-com:office:office#Keywords like '%"+Trim(oldVal)+"%'");

    var itsCounter=its.Count;
    	if (itsCounter>0)
			{
				for (var x=itsCounter;x>0;x--)
				{
				try{
        var e=its(x).itemProperties.item(catProperty).Value;
				e=e.replace(new RegExp(",","g"),";");
				var t=e.split(";");
				var nca="";
					for (var y=0;y<=t.length;y++)
					{
						if (notEmpty(t[y]))
						{
						   if (trimLow(t[y])==trimLow(oldVal))
						   {nca=nca+newVal+";";}
						   else
						   {nca=nca+t[y]+";";}
						}
					}
				its(x).itemProperties.item(catProperty).Value=nca;	
				its(x).Save();
				}catch(e){}
        }
			}
			
}

function getActionTagName()
{
//returns the active next action tag name
var ix=tagStore.find("id",new RegExp("^6$"));
var r=tagStore.getAt(ix);
try{return r.get("tag");}catch(e){return null;}
}

function getWaitTagName()
{
//returns the active next action tag name
var ix=tagStore.find("id",new RegExp("^8$"));
var r=tagStore.getAt(ix);
return r.get("tag");
} 

function deleteTag(id,node,grid)
{
//delete a tag, move child tags to root, remove all those tags from actions, remove node, remove grid entry
var tagName=getTag(id);
if (getTagSys(id)==true){alert(txtMsgDeleteSysNo);return;}

var istag=getTagType(id);
//find the tag and delete it
var ix=tagStore.find("id",new RegExp("^"+id+"$"));
var r=tagStore.getAt(ix);
var tnr=r.get("tag");
//updateSync(r,id);
tagStore.remove(r);


//update outlook category list (ol07)
updateOLCategory(tnr,null,true);

  if (node!=null)
  {//if delete was called from a tree...
  node.remove();
  gridRemoveTag(id);
  } 
  if (grid!=null)
  {
  //if delete was called from grid...
  try{Ext.getCmp("tree").getNodeById(id).remove();}catch(e){}
  }

//change children parent to 0
tagStore.filter("parent",new RegExp("^"+id+"$"));
for (var x=0;x<tagStore.getCount();x++)
{
 var r=tagStore.getAt(x);
  r.beginEdit();
  r.set("parent",0);
  r.endEdit();
if (Ext.getCmp("tvpanel").body.dom.innerText!='')
{
   var nd=new Ext.tree.TreeNode({id:r.get("id"),text:r.get("tag"),draggable:true,expandable:true,icon:getNodeIcon(r),listeners:{expand:expNode}});
   Ext.getCmp("tree").getNodeById("tags").appendChild(nd); 
}
}
tagStore.clearFilter();
//save changes
  syncStore(tagStore,"jello.tags");
  jese.saveCurrent();    
//if it was a tag remove instance from actions
if (istag){replaceItemTag(tagName,"");}
}


function gridRemoveTag(id)
{
//test if action grid is on and tag is there to delete it
try{
var s=Ext.getCmp("grid").getStore();
s.filter("entryID",new RegExp("^"+id+"$"));
for (var x=0;x<s.getCount();x++)
{
 var r=s.getAt(x);
 var tp=r.get("type");
 if (tp=="t"){s.remove(r);}
}
s.clearFilter();
}catch(e){}
}

function createTagFromCombo(tg,area)
{
//create a non existing tag from action combo
if (isValidTag(tg)==false){return;}
var choice=false;
if (choice=confirm(txtMsgNewTag.replace("%1",tg))==false){return null;}

var ttr=new tagComboR({tag:tg});
Ext.getCmp("tags").store.addSorted(ttr);
createTag(tg);
return tg;
}

function createTag(tg,noTree,par,isFolder,nomsg,isProject)
{
if (isValidTag(tg)==false){return false;}
if (!notEmpty(par)){par=0;}
if (!notEmpty(isFolder)){isFolder=false;}
jello.lastTagId++;
var tr=new tagRecord({
id:jello.lastTagId,
tag:tg,
parent:par, 
istag:!isFolder,
isproject:isProject,
modified:new Date(),
notes:"",
isprivate:false
});
var ar=new actionObject(tr,true);
var newRec=new actionRecord(ar);
tagStore.add(tr);

syncStore(tagStore,"jello.tags");
jese.saveCurrent();
getTagsArray();

//reload existing combos
try{Ext.getCmp("itags").store.loadData(globalTags);Ext.getCmp("itags").collapse();}catch(e){}
try{Ext.getCmp("tags").store.loadData(globalTags);Ext.getCmp("tags").collapse();}catch(e){}
for (var x=0;x<20;x++)
{
var olitg="itags:"+x;
try{Ext.getCmp(olitg).store.loadData(globalTags);Ext.getCmp(olitg).collapse();}catch(e){}
}

updateOLCategory(tg);
try{if (Ext.getCmp("tvpanel").body.dom.innerText==''){noTree=true;}}catch(e){}

try{
if (noTree!=true)
{
//update tree
var nd=Ext.getCmp("tree").getNodeById("tags");
	if (nd.firstChild!=null)
	{
	var nnd=new Ext.tree.TreeNode({text:tg,id:jello.lastTagId,expandable:true,icon:getNodeIcon2(true),expanded:false});
	nd.appendChild(nnd);
	}
}}catch(e){}
try{if (nomsg!=true){Ext.info.msg(txtNewTagTtl, txtTag+' '+tg+' '+txtNewTagSucc);}}catch(e){}
return jello.lastTagId; 
}

function isValidTag(s,noMsg)
{//check new tag name for invalid characters
if (s.charCodeAt(0)==32){alert("Tag name cannot begin with a space");return false;};
try{var iv=true;
	for (var x=0;x<=(s.length-1);x++)
	{
	var o=s.charCodeAt(x);
	if (o==38 || o==39 || o==126 || o==44 || /* + o==43 ||*/ o==59 || o==42 /* brackets || o==91 || o==93*/){if (noMsg!=true){alert(txtMsgChar+" ( "+s.charAt(x)+ " )");}iv=false;}
	}
}catch(e){iv=false;}
return iv;
}

function selectTagFolder()
{
//tag folder selection dialog
var t=NSpace.PickFolder();
	if (t!=null )
	{
	if (t.EntryID==jello.inboxFolder){alert(txtMsgNoSystemFolder);return;}
  if (t.DefaultItemType==0)
  {
  //must be a mail folder
  Ext.getCmp("farchiveID").setValue(t.EntryID);
  Ext.getCmp("farchive").setValue(t.FolderPath);
  }
  else
  {
  alert(txtPromptMailFolder);
  }
}

}

function getUnknownTag(utg,id)
{
//convert an unknown tag (outlook category) to a jello tag
try{
var rp=Ext.getCmp("reviewpanel").getEl().getHeight();
alert(txtNoCreateUnknTag);return;}catch(e){}
var ntg=utg.replace(new RegExp("~","g")," ");
var choice=confirm(txtMsgUnknTag.replace("%1",ntg));
if (choice==false){return choice;}
createTag(ntg);
refreshWholeView=true;
  if (id!=null)
  {
  if (refreshWholeView){refreshView();}
  }
  else
  {
  return choice;
  }
}

function checkForNotExistingTags(cats)
{
//check a category series for non existing tags. Ask to create
cats=cats.replace(new RegExp(",","g"),";");
var c=cats.split(";");
var ret="";
  if (c.length>0)
  {
    for (var x=0;x<c.length;x++)
    {
    var t=Trim(c[x]);
    var tg=getTag(t);
      if (tg==null)
      {//non existing tag. Check with user
      var choice=getUnknownTag(t);
      if (choice){ret+=t+";";}
      }
      else
      {
      ret+=t+";";
      }
    } 
  }
return ret;
}

function getTagParent(tid){
tid=Ext.escapeRe(tid);
var ix=tagStore.find("tag",new RegExp("^"+tid+"$"));
var r=tagStore.getAt(ix);
if (r!=null){return r.get("parent");}
else{return 0;}
}

function tagManager()
{//screen to manage all tags

getTagsArray();
initScreen(true,"tagManager()");
thisGrid="ggrid";
var cc=tagStore.getCount();
updateCounter(cc);

var gridTBar=TagToolbar();
sm= new Ext.grid.CheckboxSelectionModel({});
    var grid = new Ext.grid.GridPanel({
        store: tagStore,
        id:'ggrid',
        tbar:gridTBar,
        columns: [
        sm,
			{header: "", width: 30, hideable:false, fixed:true, sortable: false, renderer: renderTg, dataIndex: 'icon'},
              {header: txtTagName, width: 300, renderer:gridTagName, sortable: true,dataIndex: 'tag'},
            {header: 'ID', width: 50, sortable: true, hidden:true,dataIndex: 'id'},
            {header: txtParent, width: 50, sortable: true,hidden:true, dataIndex: 'parent'},
            {header: txtInactive, width: 50, sortable: true,renderer:tgcbox, dataIndex: 'archived'},
			{header: txtCompleted, width: 50, sortable: true, renderer:tgcbox, dataIndex: 'closed'},
            {header: txtMarkContextPrivate, width: 50, sortable: true, renderer:tgcbox,dataIndex: 'private'},
            {header: txtSystem, width: 50, sortable: true, renderer:tgcbox,dataIndex: 'sys'},
            {header: txtFolder, width: 50, sortable: true, renderer:tgcboxNo,dataIndex: 'istag'},
            {header: txtProject, width: 50, sortable: true, renderer:tgcbox,dataIndex: 'isproject'},
            {header: txtJArcFol, width: 140, sortable: true, renderer:tagManFolder, dataIndex: 'archive'},
            {header: 'Modified', width: 140,hidden:true, sortable: true, dataIndex: 'modified'},
            {header: txtNotes, width: 100,hidden:true, sortable: true, dataIndex: 'body'}
        ],
        stripeRows: true,
        autoScroll:true,
        enableColumnHide:true,
        split: true,
        deferRowRender:false,
        viewConfig:{
        emptyText:txtGridNoData
        },
        region: 'north',
        trackMouseOver:true,
        height:200,
		sm:sm,
    listeners:{
    mouseover: function(e){thisGrid='ggrid';},
 		
    rowdblclick: function(g,row,e){tagSelected(Ext.getCmp("aedit"));}
     
         },
        enableColumnMove:true,
        border:true,
        layout:'fit'
        
    });

    var ppnl=Ext.getCmp("portalpanel");
    ppnl.add(grid);
    ppnl.setAutoScroll(false);
    ppnl.doLayout();
   
   
       	grid.getSelectionModel().on('rowselect', function(sm, rowIdx, r) {
		var detailPanel = Ext.getCmp('previewpanel');
		actionTpl.overwrite(detailPanel.body, r.data);
			
	   });

   
   
    grid.on('columnresize',function(index,size){saveGridState("ggrid");});
grid.on('sortchange',function(){saveGridState("ggrid");});
grid.getColumnModel().on('columnmoved',function(){saveGridState("ggrid");});
grid.getColumnModel().on('hiddenchange',function(index,size){saveGridState("ggrid");});
try{restoreGridState("ggrid");}catch(e){}


    setTimeout(function(){
    resizeGrids();
    resizeSidebar(Ext.getCmp("west-panel"));
try{
if (jello.selectFirstItem==1 || jello.selectFirstItem=="1"){grid.getSelectionModel().selectRow(0);}
grid.getView().focusRow(0);grid.focus();}catch(e){}
	},10);

ppnl.setTitle("<img src=img\\tags16.png style=float:left;>Tag Manager");
}

function renderTg(v,m,r)
{
//render the tag name in grid

  var isp=r.get("isproject");
  var ist=r.get("istag");
  var prv=r.get("private");
  var arv=r.get("archived");
  var cls=r.get("closed");
  
  if (!ist){ic="folder.gif";if(prv){ic="folder_lock.gif";}}
if (ist){ic="list_components.gif";if(prv){ic="page_lock.gif";}}
if (isp){ic="project.gif";}
if(arv){ic="box.gif";}
if (prv==null){ic="page_lock.gif";}
if(cls){ic="check.gif";}
var rt="<img src=img\\"+ic+">";
  return rt;//+"<b>"+v+"</b>";
 
}

function tgcbox(v,m,r)
{
//render checkmarks
var rt="&nbsp;";
if (v){rt="<img src=img\\check.gif>";}
return rt;
}

function tgcboxNo(v,m,r)
{
//render checkmarks opposite
var rt="&nbsp;";
if (!v){rt="<img src=img\\check.gif>";}
return rt;
}

function TagToolbar()
{//render the tag manager toolbar

    tbar1 = new Ext.Toolbar();
        tbar1.add(
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
			editTag(null,false);
            }
            }
        }  
        ,'-',{
        icon: 'img\\page_edit.gif',
        cls:'x-btn-icon',
        tooltip: txtEdSelItmInf,
        id:'aedit',
        handler : tagSelected
        },
        {
        icon: 'img\\page_delete.gif',
        cls:'x-btn-icon',
        tooltip: txtDelItmInfo,
        id:'adel',
        handler : tagSelected        
        },'-',{
                tooltip: txtStartStop,
                cls:'x-btn-icon',
                id:'astart',
                icon: 'img\\widget-add.gif',
                handler:tagSelected        
        },{
                tooltip: txtSetPrivacy,
                cls:'x-btn-icon',
                id:'alock',
                icon: 'img\\page_lock.gif',
                handler:tagSelected        
        },{
                tooltip: txtTaskCnvToPrj,
                cls:'x-btn-icon',
                id:'aprjcon',
                icon: 'img\\project.gif',
                handler:tagSelected        
        },'-',{
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
                store:tFilterValues,
                hideTrigger:false,
                valueField:'value',
                displayField:'text',
                width:150,
                listWidth:150,
                triggerAction:'all',
                selectOnFocus:true,
				editable:false,
                emptyText:txtFilter+'...',
                mode:'local',
                listeners:{
                specialkey: function(f, e){
				if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                filterActions();}},
				select: function(cb,rec,idx){
	            TagFilterActions();
		        }
                }
                });
tbar1.insert(8,fltCmb);
      return tbar1;

}

function TagFilterActions()
{
//tag filter combo actions
var f=Ext.getCmp("filter").getValue();
tagStore.clearFilter();
if (f==5)
{
//show tags
//tagStore.filter([{property:'istag',value:'true'},{property:'isproject',value:'false'}]);
tagStore.filterBy(function ff(r){
var a=r.get("istag");
var b=r.get("isproject");
if (a==true && b==false){return true;}
});
}
if (f==1)
{
//show projects
tagStore.filter("isproject",true);
}
if (f==2)
{
//show folders
tagStore.filter("istag",false);
}
if (f==3)
{
//show inactive
tagStore.filter("archived",true);
}
if (f==4)
{
//show private
tagStore.filter("private",true);
}

}

function tagSelected(btn,vl)
{//tag manager toolbar actions
if (btn.id=="adel")
{
if (choice=confirm(txtMsgDelItem)==false){return;}
}

var sr=Ext.getCmp("ggrid").getSelectionModel().getSelections();


	for (var x=0;x<sr.length;x++)
	{
	 var t=sr[x].get("id");
	 if (btn.id=="aedit" || btn=="aedit"){editTag(t);return;}
	 
   if (btn.id=="adel"){deleteTag(t,null,Ext.getCmp("ggrid"));}
   
   if (btn.id=="astart"){startStopAction(t);}
   
   if (btn.id=="alock"){setTagPrivacy(t);}
   
   if (btn.id=="aprjcon")
   {
	var r=getTagByID(t);
		if (r!=null)
		{
		var pj=r.get("isproject");
		r.beginEdit();
		r.set("isproject",!pj);
		r.endEdit();
		syncStore(tagStore,"jello.tags");
		jese.saveCurrent();

		}
   }
	
	}

}

function getTagSubs(id)
{
tagStore.clearFilter();
tagStore.filter("parent",new RegExp("^"+id+"$"));
var re=tagStore.getCount();
tagStore.clearFilter();
return re;
}

function TagFullPath(parent,ctx,justPath)
{
//return the full path in hierarchy of a tag in HTML
if (ctx=="##"+txtInbox){return ctx;}

if (justPath==null || justPath){ret="/"+Ext.util.Format.htmlEncode(ctx);}else{ret="/"+ctx;}
if (!notEmpty(parent)){parent=0;}
if (ctx==txtJObject && parent==0){ret=ctx;return ret;}
  if (parent==0 || parent=="0")
  {
  if (justPath==null || justPath){ret="<a class=jellopathlink onclick='showTagFolder(null,null);'>"+txtJObject+"</a>"+ret;return ret;}
  }
  
  while(parent!=0)
  {
  var r=getTagByID(parent);  
      if (notEmpty(r))
      {
      var ptag=r.get("tag");
      parent=r.get("parent");
      var pid=r.get("id");
        if (justPath==null || justPath)
        {ret="/"+"<a class=jellopathlink onclick='showTagFolder(null,"+pid+");'>"+Ext.util.Format.htmlEncode(ptag)+"</a>"+ret;}
        else
        {ret="/"+Ext.util.Format.htmlEncode(ptag)+ret;}
      
      }
      else{parent=0;}
  }
 
 if (justPath==null || justPath){ret="<a class=jellopathlink onclick='showTagFolder(null,null);'>"+txtJObject+"</a>"+ret;}
 if (!justPath && ret.substr(0,1)=="/"){ret=ret.substr(1,ret.length);}
 if (!justPath && ret.substr(0,1)=="/"){ret=ret.substr(1,ret.length);}

 return ret;
}


function tagManFolder(v)
{
//return name of OL folder
var of="";
try{of=NSpace.GetFolderFromID(v);}catch(e){}
return of;
}

function gridTagName(v,m,r)
{
//tag name renderer
var hv=Ext.util.Format.htmlEncode(v);
return hv;
}

function getTagFilter(tid)
{
 	var fil="0";
 	var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
	var r=tagStore.getAt(ix);
	if (r!=null)
	{
	 try{fil=r.get("filter");}catch(e){fil="0";}
	}
	tagStore.clearFilter();
	if (!notEmpty(fil)){fil="0";}
	return fil;
}

function setTagFilter(tid,filter)
{
	try{
	var ix=tagStore.find("id",new RegExp("^"+tid+"$"));
	var r=tagStore.getAt(ix);
	if (r!=null)
	{
	var tname=r.get("tag");
		r.beginEdit();
		r.set("filter",filter);
		r.endEdit();
	  tagStore.clearFilter();
	  syncStore(tagStore,"jello.tags");
	  jese.saveCurrent();
	  Ext.info.msg(txtFilter,txtifoFilter.replace("%1",tname));
  
	}
	}catch(e){}
}

function customIconTagForm(fid)
{
//return custom icon control

  var ix=tagStore.find("id",new RegExp("^"+fid+"$"));
  var r=tagStore.getAt(ix);
  var ii=r.get("icon");
  var it=r.get("tag");
  var isf=r.get("istag");
  pLatest();
var ret="<p align=center><table>";
var fname="ud-tags";var inum="54";var irow=1;
if (!isf){fname="ud-folder";inum="27";}
  for (var x=1;x<=inum;x++)
  {
  if (irow==1){ret+="<tr>";}
  var sel="";
  if (ii==fname+" ("+x+").gif"){sel="style='border:1px double black;background:yellow;'";}
  ret+="<td width=24 height=24 align=center valign=middle "+sel+"><a class=jellolinktop onclick='setTagIcon("+fid+",&quot;"+fname+" ("+x+")&quot;);'><img "+sel+" src='img\\"+fname+" ("+x+").gif'></td><td>&nbsp;</td>";
  irow++;
  if (irow==10){ret+="</tr>";irow=1;}  
  }
  ret+="</table>";
  if (!notEmpty(ii)){ret+="<br>Currently Using the default icon";}
  else
  {ret+="<br><a class=jellolinktop onclick='setTagIcon("+fid+",null)'>Reset to the default icon</a>";}
  ret+="</p>";
Ext.getCmp("latestform").update(ret);
Ext.getCmp("latestform").setTitle(it);
Ext.getCmp("thelatestform").setHeight(300);
Ext.getCmp("thelatestform").setTitle("Select Icon");


}

function setTagIcon(fid,icon)
{
//set custom icon to tag

   var ix=tagStore.find("id",new RegExp("^"+fid+"$"));
   var r=tagStore.getAt(ix);

  r.beginEdit();
  if (notEmpty(icon)){r.set("icon",icon+".gif");}else{r.set("icon",null);}
  r.endEdit();
  tagStore.clearFilter();
  syncStore(tagStore,"jello.tags");
  jese.saveCurrent();
 var nd=Ext.getCmp("tree").getSelectionModel().getSelectedNode();
 if (notEmpty(icon))
 {nd.setIcon("img\\"+icon+".gif");}
 else
 {nd.setIcon(getNodeIcon(r));}
 Ext.getCmp("thelatestform").destroy();
 customIconTagForm(fid);
}

function addJournalInTag(subj,otherTag)
 {

 try{var jf=NSpace.GetFolderFromID(jello.nonActionFolder);}catch(e)
 {alert(txtWdImpNoJou);return;}

 var su=subj;
 if (!notEmpty(su)){su=prompt(txtJournalPrompt);}
 if (!notEmpty(su)){return;}
  var jit=jf.Items.Add();
 jit.Subject=su;
 jit.Start=DisplayDate(new Date());
 if (otherTag==null){jit.itemProperties.item(catProperty).Value=lastContext;}
 else{jit.itemProperties.item(catProperty).Value=otherTag;}

 try{jit.Type="Note";}catch(e){}
 jit.Save();

if (otherTag==null)
	 {
		var ar=new actionObject(jit,true);
	  	var newRec=new actionRecord(ar);
		var g=getActiveGrid();
		var store=g.getStore();
		store.insert(0,newRec);
		g.getSelectionModel().selectFirstRow();
		try { updateAFooterCounter(store);}catch(e){}
	 }
	 else
	 {return jit;}
 }

  function addReferenceItems(ctx,listStore,subj)
 {
 //query reference items for this tag
 try{var jf=NSpace.GetFolderFromID(jello.nonActionFolder).Items;}catch(e){return;}

 ctx=ctx.replace(new RegExp("~","g")," ");
  var dasl="urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "'";
 var its=jf.Restrict("@SQL="+dasl);
 var count=its.Count;
	if (count==0){return;}
	for (var x=1;x<=count;x++)
	{
	var jjj=its(x);
	var ar=new actionObject(jjj,true);
  	var newRec=new actionRecord(ar);
    listStore.add(newRec);
	}

 }

function printProjectHistory()
{
//print project history
pjid=lastOpenTagID;
var pj=getTagByID(pjid);
var pjname=pj.get("tag");
var ret="<head><META content='text/html; charset=utf-8' http-equiv=Content-Type><style>.printChars{height:30px;font-size: 14px;font-family: 'Segoe UI', Verdana, 'Trebuchet MS', Arial, Sans-serif;border-bottom: 1px solid Gainsboro;}</style></head><body>";
var jou="<tr><td colspan=10 class=printChars><br><b>Reference/Journal Items</b></td></tr>";
ret+="<h2>"+pjname+"</h2><p>"+pj.get("notes")+"</p>";
ret+="<table width=98%><thead class=printChars><tr><td><b>Created</b></td><td><b>Subject</b></td><td><b>Completed</b></td></tr></thead>";
//get all items
var pjitems=showContext(pjname,2,null,4);
addReferenceItems(pjname,pjitems);

var renderPJItem= function(r){
var cl=r.get("iclass");
	if (cl==42)
	{

		try{var it=NSpace.GetItemFromID(r.get("entryID"));
		var nts=it.Body;}catch(e){}
	jou+="<tr><td valign=top class=printChars colspan=10>"+r.get("subject")+"<br><span style='font-size:9px'>"+nts+"</span></td>";
	}
	else
	{
	var nts=r.get("body");
	var s=nts.search(txtTryToGetBody);
	if (s>-1){nts=null;}
	if (!notEmpty(nts)){nts=r.get("notes");}else{nts="";}
	ret+="<tr><td valign=top class=printChars width=10%>"+DisplayDate(r.get("created"))+"</td>";
	ret+="<td valign=top class=printChars width=80%>"+r.get("subject")+"<br><span style='font-size:9px'>"+nts+"</span></td>";
	ret+="<td valign=top class=printChars width=10%>"+DisplayDate(r.get("completed"))+"</td>";
	ret+="</tr>";
	}
};

pjitems.sort("created","ASC");
pjitems.each(renderPJItem);
ret+=jou+"</table></body></html>";

//create post
var ff=NSpace.GetFolderFromID(jello.inboxFolder).Items;
var it=ff.Add(6);
it.Subject="Project History "+pjname;
it.HTMLBody=ret;
it.Display();
}
// show all tasks that have any of the cats of the selected item
function tasksWithTag(rec, tagmatch)
{

	var dasl = "";
    var cats = "";
	var clist = new Array();
	if( rec != null)
        cats = rec.get("tags");
    else{
		tagmatch=tagmatch.replace(/~/g," ");
        cats = tagmatch;
    }
    if( cats == "" || cats == null || typeof(cats) =="undefined" ){
        Ext.Msg.alert("","No items found");
        return;
    }
	if( cats.indexOf(",") == -1){
		clist.push(cats);
	}else{
		clist = cats.split(",");
	}
	for( var i = 0; i < clist.length; i++){
		var ctx = clist[i];
		if (dasl != "") dasl += " OR ";
		dasl +="(urn:schemas-microsoft-com:office:office#Keywords LIKE '" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + ";%' OR urn:schemas-microsoft-com:office:office#Keywords LIKE '%;" + ctx + "' OR urn:schemas-microsoft-com:office:office#Keywords = '" + ctx + "')";
	}
	var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
	var its=iF.Restrict("@SQL="+dasl);
    if( its.Count == 0){
        Ext.Msg.alert("","No items found");
    }else
        buildGrid(null,its,0,0,null,true,false);

}

function addTagToAL(tagname,olitem,rec,noRec)
{//add a tag to an actionlist item
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
//updateInboxRecord(rec);
	if (wasNew)
	{
	setTimeout(function(){refreshView();},6);
	}
}

function  goToTagQuestion(tg)
{
if (confirm(txtGoto+" "+tg+" ?")==true){
           showContext(tg,1);
}
}

function getUniqueTags(tags){
//j5
//get tags from multiple items and return only the unique ones
if (tags==null || tags=="" || typeof(tags)=="undefined"){return;}
var e=tags.replace(new RegExp(",","g"),";");
var tag=e.split(";"); var ret="";
var newtags=new Array();
	for (var x=0;x<tag.length;x++)
	{
		var exists=false;
		if (x==0){}
		var y=0;
		do
		{
		if (trimLow(tag[x])==trimLow(newtags[y])){exists=true;break;}
		y++;
		}while(y<=newtags.length);
		if (exists==false){newtags.push(tag[x]);}
	}
	var ret="";
for (var y=0;y<newtags.length;y++)
		{
		ret+=newtags[y]+";";
		}
		return ret;
}
function tagList(tagdisp,ctx,tags,id){
//j5
//create the tag list for an item


	
var rt=new Array();var isnext=false;
if (tags==null || tags=="" || typeof(tags)=="undefined"){rt.push(" ");rt.push(false);return rt;}
var tge;

//Jeremy's mod
 try
    {
          tge=tags.replace(new RegExp(",","g"),";");
    }
 catch(e)
    {
          //tge="";
          tags+='';
          tge=tags;
    }
//----
var tag=tge.split(";"); var ret="";
tag.sort();
	for (var x=0;x<tag.length;x++)
	{
		if (notEmpty(tag[x]))
		{
		var oo=Trim(tag[x]);

	oo=Ext.util.Format.htmlEncode(oo);
	if (isValidTag(oo,true))
	{
		tag[x]="<span class=notag title='"+oo+"'>"+oo+"</span>";
		var ooo=oo.replace(new RegExp(" ","g"),"~");
      //if (ctx==null || ctx=="[INBOX]"){var ixc="style=cursor:auto;";var ix="<span class=tagremove onclick=removeTag('"+ooo+"','"+id+"');>x</span>&nbsp;";}else{var ix="";var ixc="";}
      var ixc="style=cursor:hand;";var ix="<span class=tagremove onclick=removeTag('"+ooo+"','"+id+"','"+tagdisp+"');>"+elmRemoveTag+"</span>&nbsp;";//}else{var ix="";var ixc="";
      var tagClass="tagcon";
      var theTag=getTag(oo);
      if (theTag==null){tagClass="notag";ix="";ixc="style=cursor:hand; title='"+txtConvtoJTag+"' onclick=getUnknownTag('"+ooo+"','"+id+"');";}
      else
      {if (theTag.get("isproject")){tagClass="tagprj";}
      }

      var tconame=oo;
  		if (tconame.length>16){tconame=tconame.substr(0,16)+"...";}
  		tag[x]="<a title='"+oo+"' onclick=goToTagQuestion('"+ooo+"'); class="+tagClass+" "+ixc+">"+tconame+"</a>"+ix;
			if (trimLow(oo)!=trimLow(getActionTagName()))
			{
				if (trimLow(ctx)!=trimLow(oo))
				{ret+=tag[x];}
			}
			else{isnext=true;}
	//		if (id==null){ret+=tag[x];}
	}
	else
	{
	//mark as having invalid category characters
	ret="<span class=tagprj>"+txtInvalidCharsCat+"</span>";
	}

		}

    }
if (ctx=="[INBOX]"){ctx=null;}
if (ctx==null && isnext==true){ret+="<a class=tagnext style=cursor:auto;>"+txtNextAction+"</a><span class=tagremove onclick=removeTag('"+getActionTagName()+"','"+id+"','"+tagdisp+"');>"+elmRemoveTag+"</span>&nbsp;";}
rt.push(ret);
rt.push(isnext);

return rt;
}

function quickAddTag(tg, tgfield, tagdisp, id){
//j5
//quick add a tag from combobox into action form

var vl=Ext.getCmp(tg).getRawValue();
  if (notEmpty(vl))
  {//if there is no such tag create it
  if (getTag(vl)==null){vl=createTagFromCombo(vl,tg);}
  }
  if (!notEmpty(vl)){Ext.getCmp(tg).setValue("");return;}

var nv=Ext.getCmp(tgfield).getValue();
var existsIn=false;
var e=nv.replace(new RegExp(",","g"),";");
var tag=e.split(";");
for (var x=0;x<tag.length;x++)
{
if (trimLow(vl)==trimLow(tag[x])){existsIn=true;}
status="Tagging...";
}
if (existsIn==false){Ext.getCmp(tgfield).setValue(nv + ";"+vl);}

var f=tagList(tagdisp,null,Ext.getCmp(tgfield).getValue(), id);
var s=f[0];
Ext.getCmp(tagdisp).getEl().dom.innerHTML=s;
Ext.getCmp(tg).reset();

}

function removeTag(tg,id,tagdisp){
//remove tag from form's X
tg=tg.replace(new RegExp("~","g")," ");
if (id!="undefined" && id != "000")
{removeInboxTag(tg,id, tagdisp);}

var newval="";
	try{
	var nv=Ext.getCmp('tagfield');
	var tags=nv.getValue();
	}catch(e)
	{
	//action list
	var it=NSpace.GetItemFromID(id);
	var tags=it.itemProperties.item(catProperty).Value;
	var newval="{{000}}";
	}
var e=tags.replace(new RegExp(",","g"),";");
var tag=e.split(";");
for (var x=0;x<tag.length;x++)
{
if (trimLow(tg)!=trimLow(tag[x])){newval+=tag[x]+";";}
}

	if (newval!="{{000}}")
	{
      	try{
        nv.setValue(newval);
      	var f=tagList("formtagdisplay",null,nv.getValue());
      	var s=f[0];
      	Ext.getCmp(tagdisp).getEl().dom.innerHTML=s;
      	}catch(e){}
	}
	else
	{


	}

}

function showTagFolder(r,rid)
{
//show folder's contents
if (rid!=null)
{
  var ix=tagStore.find("id",new RegExp("^"+rid+"$"));
  r=tagStore.getAt(ix);
}
if (r!=null){
var fname=r.get("tag");
var fid=r.get("id");
var ftag=r.get("istag");
}else{
var fname=txtJObject;
var fid=0;
var ftag=false;
}
var far=new Array();
tagStore.filter("parent",new RegExp("^"+fid+"$"));
for (var x=0;x<tagStore.getCount();x++)
{
 var r=tagStore.getAt(x);
 far.push(r);
}
tagStore.clearFilter();
lastOpenTagID=fid;

showContext(fid,1,null,null,far,ftag);

}


function tagUpLevel(parentID)
{
//go up one level in actionlists
	var ix=tagStore.find("id",new RegExp("^"+parentID+"$"));
	var r=tagStore.getAt(ix);
	try{showTagFolder(r);}catch(e){alert(e);}
}
