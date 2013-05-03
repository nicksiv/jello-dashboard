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


//Contacts & attachments


var contactList=new Array();
var conToRecordsList=new Array();

function delegateItems(recs,fromForm)
{
//delegate item to contact, add to @waiting and archive
conToRecordsList=recs;
try{cwin.destroy();}catch(e){}
try{conp.destroy();}catch(e){}
try{Ext.getCmp("thecontactform").destroy();}catch(e){}

 var conp = new Ext.FormPanel({
        labelWidth: 75,
        frame:true,
        title: txtDelgToCon,
        height:200,
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        id:'contactform',
        iconCls:'conformicon',
        buttonAlign:'center',
        defaults: {width: 160},
        defaultType: 'textfield',

        items: [{
                fieldLabel: txtFind,
                id: 's-con',
                value:'',
                disabled:false,
                tabIndex:0,
                emptyText:txtDelgToConCri,
                allowBlank:true,
                listeners:{
		    specialkey: function(f, e){
            if(e.getKey()==e.ENTER){
                userTriggered = true;
                e.stopEvent();
                f.el.blur();
                searchContacts();}}
                }
            },  new Ext.form.Label({
				html:'<table width=100%><tr><td class=formconlistchars><i>"+txtDelgToConCri2+"</i></td></tr></table>',
				id:'conview',
				cls:'formconlist',
				layout:'anchor',
				scrollable:true,
				autoScroll:true,
				autoWidth:true
            })
        ]

    });

 var cwin = new Ext.Window({
        title: txtContacts,
        width: 450,
        height:300,
        id:'thecontactform',
        resizable:false,
        draggable:true,
        layout: 'fit',
        plain:false,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        items: conp,
        modal:true,
        listeners: {destroy:function(t){showOLViewCtl(true);}},
        buttons: [{
            text:txtSave,
            id:'s-but',
            tooltip:txtSet+' [F2]',
            listeners:{
            click: function(b,e){
            if (addContactsToItems(conToRecordsList,contactList)!=true)
            {cwin.destroy();}
            }
            }
        },{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
                        listeners:{
            click: function(b,e){

            cwin.destroy();}
            }
        }]
    });



    showOLViewCtl(false);
    cwin.show();
    cwin.setActive(true);

  updateContactList();
  setTimeout(function(){Ext.getCmp("s-con").focus(true,100);},1);

//shortcut keys
var map = new Ext.KeyMap('thecontactform', {
    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
    try{Ext.getCmp("thecontactform").destroy();}catch(e){}
    try{Ext.getCmp("theactionform").setActive(true);}catch(e){}
    },
    scope: this});

var map2 = new Ext.KeyMap('thecontactform', {
    key: Ext.EventObject.F2,
    stopEvent:true,
    fn: function(){
    addContactsToItems(conToRecordsList,contactList);
    cwin.destroy();
    },
    scope: this});

}

function searchContacts()
{
//perform a search for a contact
//contactList array must be used for storing contacts
var sc=Ext.getCmp("s-con").getRawValue();
if (notEmpty(sc)==false){return;}
var its=NSpace.GetFolderFromID(jello.contactFolder).Items;
if (sc!="*")
{
var dasl="(urn:schemas:contacts:cn LIKE '%"+sc+"%')";
for (var i =1; i <= 3; i++)
	   dasl += " OR (urn:schemas:contacts:email" + i+ " LIKE '%"+sc+"%')";
dasl = "("+dasl+")";
var iits=its.Restrict("@SQL="+dasl);
}
else
{iits=its;}

var ret="<tr><td colspan=2 class=formconlistchars><b>"+txtContactsFound+"</b></td></tr>";
for (var x=1;x<=iits.Count;x++)
{
ret+="<tr><td width=20><a class=jellolink onclick=getContact('"+iits(x).EntryID+"') title='"+txtConSelect+"'><img src=img\\page_new.gif></a></td><td><a class=jellolink onclick=olItem('"+iits(x).EntryID+"') title='"+txtConView+"'>"+iits(x).fileas+"</a></td></tr>";
}

if (iits.Count==0){ret="<tr><td colspan=2 class=formconlistchars><i>"+txtConNoMatch+"</i></td></tr>";}
updateContactList(ret);
Ext.getCmp("s-con").focus(true);
}

function getContact(id)
{
//select a contact from contact selector form
contactList.push(id);
updateContactList();
}

function updateContactList(info)
{
var ret="<table width=100%><tr><td colspan=2 class=formconlistchars><b>"+txtContactsSelected+"</b></td></tr>";
	for (var x=0;x<contactList.length;x++)
	{
	var it=NSpace.GetItemFromID(contactList[x]);
	ret+="<tr><td width=20><a class=jellolink onclick=delContact("+x+") title='"+txtConRmv+"'><img src=img\\delete.gif></a></td><td><a class=jellolink onclick=olItem('"+it.EntryID+"') title='"+txtConView+"'>"+it.fileas+"</a></td></tr>";
	}
ret+="<tr><td colspan=2 class=formconlistchars><i>"+txtCreate+" <a class=jellolinktop onclick=createContact()><b>"+txtConNew+"</b></a></i></td>";
ret+="<tr><td style='border-bottom:1px solid black;' colspan=2>&nbsp;</td></tr>";
if (notEmpty(info))
{ret+=info+"</table>";}
else
{ret+="<tr><td colspan=2 class=formconlistchars><i>"+txtConAddAnother+"</i></td>";}
ret+="</table>";
document.getElementById("conview").innerHTML=ret;
Ext.getCmp("s-con").focus(true);
}

function delContact(id)
{
//remove a contact from contact selector form
contactList.splice(id,1);
updateContactList();
}

function createContact()
{
//create new contact
var nct=NSpace.GetFolderFromID(jello.contactFolder).Items.Add();
nct.Display();
}

function addContactsToItems(recs,cons)
{
 //add contacts to items
var noRec=false; var conlist="";
if (cons.length==0){var noco=confirm(txtDelgNoCons);if (noco==false){return true;}}
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
    if (it.Class!=48)
    {  //no tasks
      //try
      //{
      var ntask=toTagID(it,r,8,noRec);
        clearOLItemLinks(ntask);
        for (var y=0;y<cons.length;y++)
         {
         var cit=NSpace.GetItemFromID(cons[y]);
         addOLItemLink(ntask,cit);
         sendDelgMail(cons,ntask,cit);
         conlist+=cit+" ";
         }

      ntask.Save();
      //}catch(e){}
    }
    else
    {
    //assign contacts to tasks
    clearOLItemLinks(it);
    for (var y=0;y<cons.length;y++)
         {
         var cit=NSpace.GetItemFromID(cons[y]);
         addOLItemLink(it,cit);
		sendDelgMail(cons,it,cit);
         conlist+=cit+" ";
         }

      it.Save();
	  updateActionForm(it);
    }

	}

Ext.info.msg(txtDelegated,txtDelegTo+" "+conlist);
}

function addOLItemLink(it,cit)
{
//add a link to an outlook item
    try{
    it.Links.Add(cit);
    it.Save();
    }catch(e){}
}

function clearOLItemLinks(it)
{
//clear all outlook item's links
try{
if (it.Links.Count==0){return;}
  for (var x=it.Links.Count;x>0;x--)
  {
  it.Links.Remove(x);
  }
    }catch(e){}
}

function updateActionForm(it)
{
//if there is an action form open, update links box
try{
var f=Ext.getCmp("actioncons");
actioncons.innerHTML=getActionAttachmentLink(it);
}catch(e){}
}

function actionEditAttachments()
{

    var fp = new Ext.FormPanel({
        //renderTo: 'fi-form',
        fileUpload: true,
        width: 500,
        frame: true,
        title: txtFlUploadFm,
        autoHeight: true,
        bodyStyle: 'padding: 10px 10px 0 10px;',
        labelWidth: 50,
        items: [{
            xtype: 'textfield',
            inputType:'file',
            id: 'form-file',
            width:370,
            emptyText: txtFlSelect,
            fieldLabel: txtFlAttach,
            buttonCfg: {
                text: '',
                iconCls: 'upload-icon'
            }
        }],
        buttons: [{
            text: txtUpload,
            tooltip:txtFlUpload+' [F2]',
            handler: function(){
            getFileToAttach(false);
            }
        },{
            text: txtLink,
            tooltip:txtFlLink,
            handler: function(){
            getFileToAttach(true);
            }
        },{
            text: txtCancel,
            tooltip:txtCancel+' [F12]',
            handler: function(){
			Ext.getCmp("theattachform").destroy();
            }
        }]
    });

   var awin = new Ext.Window({
        title: txtFlAttach,
        width: 500,
        height:150,
        id:'theattachform',
        resizable:false,
        draggable:false,
        layout: 'fit',
        plain:true,
        draggable:true,
        bodyStyle:'padding:5px;',
        buttonAlign:'center',
        modal:true,
        items: fp
    });

//show the form
    awin.show();
    awin.setActive(true);

//shortcut keys
var map2 = new Ext.KeyMap('theattachform', {
    key: Ext.EventObject.F2,
    stopEvent:true,
    fn: function(){
    try{
    getFileToAttach();
    }catch(e){}},
    scope: this});

var map = new Ext.KeyMap('theattachform', {
    key: Ext.EventObject.F12,
    stopEvent:true,
    fn: function(){
            try{Ext.getCmp("theattachform").destroy();}catch(e){}

            },
    scope: this});

}

function getFileToAttach(linkOnly)
{
//attach a file to active action
var v=Ext.getCmp("form-file").getValue();
	if (notEmpty(v))
	{
	var id=Ext.getCmp("olid").getValue();
	try{var it=NSpace.GetItemFromID(id);}catch(e){alert(txtInvalid);return;}
		try{
		if (!linkOnly){var att=it.Attachments.Add(v);}else{var att=it.Attachments.Add(v,6);}
		}catch(e){alert(txtInvalid);Ext.getCmp("form-file").setValue("");try{att.Delete();}catch(w){};return;}
	actioncons.innerHTML=getActionAttachmentLink(it);
	Ext.getCmp("theattachform").destroy();
	}

}
function buildWaitingGrid(ctx,iits,counter,filtering,oItems,istag)
{
//specific grid for the @waiting tag

var waitStore = new Ext.data.GroupingStore({
reader: reader,
sortInfo:{field: 'contacts', direction: "ASC"},
id:'waitingStore',
groupField:'contacts',
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
		   {name: 'tags'},
		   {name: 'contacts'}
        ]
    });

//add other type items
if (oItems!=null)
{
 for (var x=0;x<oItems.length;x++)
 {
	var ar=new actionObject(oItems[x],true);
	ar.contacts="("+txtObjects+")";
	var newRec=new actionRecord(ar);
	waitStore.add(newRec);
 }
}

//add actions
	if (iits!=null)
	{ try{
		if (iits.Count>0)
		{ 
			for (var x=1;x<=iits.Count;x++)
			{
				var lc=iits(x).Links.Count;
					if (lc>0)
					{
						for (var y=1;y<=lc;y++)
						{
						var ar=new actionObject(iits(x));
						ar.contacts=iits(x).Links.Item(y).Name;
						var newRec=new actionRecord(ar);
						waitStore.add(newRec);
						}
					}
					else
					{
					var ar=new actionObject(iits(x));
					ar.contacts="("+txtContactNo+")";
					var newRec=new actionRecord(ar);
					waitStore.add(newRec);
					}
			}
		} }catch(e){}
	}
waitStore.sort("contacts","ASC");
return waitStore;
}

function sendDelgMail(cons,it,cname)
{
			if (cons.length==1 && it.DelegationState==0 && jello.sendTaskRequests==1)
			{
			//on one contact assignment ask to assign task
			if (confirm(txtAskTaskRequest+" "+cname+" ?"))
			{
				it.Assign();
				it.Recipients.Add(cname);
				try{it.Display();/*it.Send();*/}catch(e){alert(txtDelgNoConMail);}
			}
			}

}
