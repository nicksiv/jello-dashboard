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
//    2008-2010 N.Sivridis http://jello-dashboard.net


//-------------------------------------

var twDefaultName="jelloTiddly.htm"


function wikiValue(a,b,c,d,e,f,g){
this.olid=a;
this.subject=b;
this.modified=c;
this.tags=d;
this.notes=e;
this.tgid=f;
this.notesfrom=g;
}

function tiddlyScreen()
{
//TiddlyWiki sync settings


    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");

  
var simple = new Ext.FormPanel({
        frame:true,
//        fileUpload: true,
//        height:380,
    	title:'TiddlyWiki export',
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:115,
        id:'twform',
        buttonAlign:'center',
//        defaults: {width: 320},
        defaultType: 'textfield',

        items: [
	       

	    {
                fieldLabel: 'Location of Tiddlywiki',
                inputType:'text',
                id: 'ttwsync',
                hidden:false,
                tabIndex:2,
		listeners:{blur:function(a){getTWFileInfo(false);}},
		width:420,
                allowBlank:true
            },
	    	new Ext.form.Label({
		html:'',
		width:450,
	    	id:'reslabel',
		cls:'formfilelist',
		fieldLabel:' ',labelSeparator:'',
		height:30
		}),

            {
                fieldLabel: 'Skip tags (comma separated)',
                id: 'twskip',
                visible:true,
                value:'',
                tabIndex:3,
		width:420,
                allowBlank:true
            },
            new Ext.form.Checkbox({
                fieldLabel: 'Skip Private Tags',
                id: 'skprv',
		checked:true
            }),
        new Ext.form.Label({
				html:'',
				width:650,
	    			id:'reslabel2',
				cls:'formfilelist',
            			height:40
				})
	    ,new Ext.Button({text:'<b>Create Tiddlywiki File</b>',id:'twgo',fieldLabel:' ',autoWidth:false,width:180,height:30,labelSeparator:'',listeners:{click:function(p){tiddlySync();}}})

        ]

    });


simple.render(mp.getEl());

    setTimeout(function(){
	getTWFileInfo(true);
    }, 250);



}

function getTWFileInfo(init)
{
//update selected folder/file information

if (init)
{
    if (!notEmpty(jello.lastTiddlyWiki) || (jello.lastTiddlyWiki=="undefined"))
	{
	    var jepath="";
	    var fso = new ActiveXObject("WScript.Shell");
	    jepath = fso.SpecialFolders("MyDocuments")+"\\"+twDefaultName;
	    jepath = jepath.replace(/\\/g, "\\");
	}
    else
	{var jepath=jello.lastTiddlyWiki;}

    Ext.getCmp("ttwsync").setValue(jepath);
}

var fso = new ActiveXObject("Scripting.FileSystemObject");
var fname=Ext.getCmp("ttwsync").getValue();
var fileinfo="<b><font color=red>A new file will be created</b>";
Ext.getCmp("twgo").setDisabled(false);
Ext.getCmp("twgo").setText("<b>Create Tiddlywiki File</b>");
    if (notEmpty(fname))
	{
	var twf="";
	try{twf=fso.GetFile(fname);fileinfo="File "+fname+" already exists.<br>Created on: "+twf.DateCreated+"\nFile Size:"+twf.Size+" bytes";Ext.getCmp("twgo").setText("<b>Replace Tiddlywiki File</b>");}catch(e){}

	}
    else{fileinfo="No file specified";Ext.getCmp("twgo").setDisabled(true);}
    Ext.getCmp("reslabel").update(fileinfo);

}

function tiddlySync()
{
//TiddlyWiki sync
var creator="Jello";


var newToFolder=getAppPath().replace(appName,"")+"res/";
var wikiFile=Ext.getCmp("ttwsync").getValue();
var tagSkip=Ext.getCmp("twskip").getValue();
var privSkip=Ext.getCmp("skprv").getValue();

var jpath=getAppPath().replace(appName,"")+"/res/";
//temp
//var jpath=getAppPath().replace("jelloapp.html","")+"/res/";

var sourceWiki=jpath+"tw.html";
var destWiki=wikiFile;

var tsaverJAR=jpath+"TiddlySaver.jar";

if (wikiFile.substr(wikiFile.length-5,5)!=".html" || wikiFile.substr(wikiFile.length-4,4)!=".htm") {wikiFile+=".html";}

var fso = new ActiveXObject("Scripting.FileSystemObject");
try{
var file=fso.getfile(wikiFile);
if (confirm("Replace existing file?")==false){return;}
}catch(e){}



//--fs object
    var fs = new ActiveXObject("Scripting.FileSystemObject");
    var twStream = fs.OpenTextFile(sourceWiki,1,false,0);
	var str=twStream.ReadAll();
	twStream.Close();
	var cj=str.search("jello template");
	//if (cj==-1){alert("This is not a Jello Dashboard generated wiki file!");return;}
    //get split point of the store area
	var e=0;var splitPoint=0;var s=0;var endData=0;
	var TWbody="";var JDdata="";
/*	if (mode==0)
	{//for update existing TW file
	e=str.indexOf('<div id="storeArea">');
	s=str.indexOf('<!--POST-STOREAREA-->');
	splitPoint=e+21;
	endData=s;
	var TWheader=str.substr(0,splitPoint);
	var TWfooter=str.substr(endData,10000000);
	JDdata=str.substr(e+21,s-(e+21));
	}*/

	//for a new tw file
	e=str.indexOf('<!--POST-STOREAREA-->');
	splitPoint=e-8;
	var TWheader=str.substr(0,splitPoint);
	var TWfooter=str.substr(splitPoint,10000000);
	
	
    var wikiRecord = Ext.data.Record.create([
    {name: 'olid'},
    {name: 'tgid'},
    {name: 'subject'},
    {name: 'modified'},
    {name: 'tags'},
    {name: 'notes'}
    ]);
    var wikiStore=new Ext.data.SimpleStore({
        fields: [
    {name: 'olid'},
    {name: 'tgid'},
    {name: 'subject'},
    {name: 'modified'},
    {name: 'tags'},
    {name: 'notes'}]
    });

	//get existing tiddlyWiki data
	
	if (notEmpty(JDdata))
	{
  JDdata=manualConvertUnicodeToUTF8(JDdata);
	var o=JDdata.split("</div>");

	var ret="";
	var otherTiddlers="";
	
		if (o.length>0)
		{
			for (var x=0;x<o.length;x++)
			{
				var line=o[x].split('"');
				
				var oid="0";var tid="0";var twbody="";var nfrom="";
		
          //get tiddler values to store
          var ttitle="";var tmodif="";var ttags="";var tolid="";var ttgid="";
          var systemFlag=false;
          for (var k=0;k<line.length;k++)
          {
              
              var s="";
              try{s=line[k].search(" title=");if(s>-1){ttitle=line[k+1];k++;}}catch(e){}
              try{s=line[k].search(" tags=");if(s>-1){ttags=line[k+1];k++;}}catch(e){}
              
          
              if (ttags.search("jello")>-1 && ttags.search("template")>-1){systemFlag=true;}
              if (ttags.search("excludeLists")>-1){systemFlag=true;}
              if (ttags.search("jello")>-1 && ttags.search("systemConfig")>-1){systemFlag=true;}
              
              if (!systemFlag)
              {
              s=line[k].search(" modified=");if(s>-1){tmodif=line[k+1];k++;}
              s=line[k].search(" olid=");if(s>-1){tolid=line[k+1];k++;}
              s=line[k].search(" tagid=");if(s>-1){ttgid=line[k+1];k++;}
              s=line[k].search(" notes=");if(s>-1){nfrom=line[k+1];k++;}

              s=line[k].search("<pre>");if(s>-1){twbody=line[k];}
    					var ls=twbody.search("<pre><\/pre>");
		    			if (ls==-1){twbody=twbody.substr(7,(twbody.length)-14);}else{twbody="";}
              }
                  
          }
    		
    		if (!systemFlag)
    		{
				var cr=new wikiValue(tolid,ttitle,tmodif,ttags,twbody,ttgid,nfrom);
				var newRec=new wikiRecord(cr);wikiStore.add(newRec);
				}
				else
				{
        //system tiddlers
        otherTiddlers+=o[x]+"</div>";
        }

			}
		}
	
	var changesFromTW=0;
	var newRecords=0;
	//loop through tw records
	for (var x=0;x<wikiStore.getCount();x++)
	{
		var r=wikiStore.getAt(x);

			var modif=r.get("modified");
		if (typeof(modif)!="undefined")	
		{
			var olid=r.get("olid");
			var tid=r.get("tgid");
			var ttl=manualConvertUTF8ToUnicode(r.get("subject"));
			var tg=manualConvertUTF8ToUnicode(r.get("tags"));
			var nf=r.get("notesfrom");
			var notes=manualConvertUTF8ToUnicode(r.get("notes"));
			var twDate=new Date();
			twDate=Date.convertFromYYYYMMDDHHMM(modif);
			var isAction=tg.search("jd.action");
			var isTg=tg.search("jd.tag");
			var isPj=tg.search("jd.project");
			var isNA=ttl.search("jd.next");
			var isDO=ttl.search("jd.done");
	//tag or project
	if ((isTg>-1 || isPj>-1) && isNA==-1 && isDO==-1 && tid==0)
	{
         
         
         var r=getTag(ttl);
         var pj=false;
         if (isPj>-1){pj=true;}
         if (r==null)
         {createTag(ttl,true,0,false,false,pj);newRecords++;}
         else
         {
            if (pj)
            {
              if (r.get("isproject")==false)
              {
              r.beginEdit();
              r.set("isproject",true);
              r.endEdit();
                syncStore(tagStore,"jello.tags");
                jese.saveCurrent();
              }
            }
         }
         
      }
			
      if (isAction>-1)
			{
			//outlook item
			if (olid=="0" || olid=="")
			{
			//create new action
			 
			 var newA=createActionOL(ttl)
			 newA.Categories=fixTWTags(tg,newA);
			 newA.Body=notes;
			 newA.Save();
			 olid=newA.EntryID;
			 setJNotesProperty(newA,notes);        
			 newRecords++;
			 ret+=ttl;
			}
  					
  			else
  			{
  			try{
				var it=NSpace.GetItemFromID(olid);
  				var olDate2=new Date(it.LastModificationTime);
  				var olTemp=new Date();
				var olDate1=olDate2.convertToYYYYMMDDHHMM();
  				var olDate=Date.convertFromYYYYMMDDHHMM(olDate1);
  				
  					if (twDate>olDate)
					{
					it.Subject=ttl;
					if (nf=="ol"){it.Body=notes;}
					else
					{setJNotesProperty(it,notes);}
					it.itemProperties.item(catProperty).Value=fixTWTags(tg,it);
					it.Save();
					changesFromTW++;
					}
				}catch(e){}
			}
		}
			
	}	
	}

	if (newRecords==0 && changesFromTW==0){alert("No changes found to update to Outlook");return;}
  if (!confirm(ret+"\n"+newRecords+" New records will be imported and "+ changesFromTW+" Records changes from Wiki")){return;}

	}
	
	var ff=NSpace.GetFolderFromID(jello.actionFolder);
	var f=ff.Items.Restrict("@SQL=http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2");
	
	for (var x=1;x<=f.Count;x++)
	{
	var tgstring=f(x).itemProperties.item(catProperty).Value+";"
	var ts=tgstring.split(";");
	var tagString="";
	//next action tag
  
	var nextTag="!Next";
  try{var r=getTagByID("6");nextTag=r.get("tag");}catch(e){}
  
	if (ts.length>0)
	{
		for (var y=0;y<ts.length;y++)
		{
			if (notEmpty(ts[y]))
			{
			var cs=Trim(ts[y]);
			if (cs==nextTag){cs="jd.next";}
			var scs=cs.search(" ");
				if (scs>-1){cs="[["+cs+"]]";}
			 tagString+=cs+" ";
			 }
		}
		if (tagString.substr(tagString.length-1,1)==" "){tagString=tagString.substr(0,tagString.length-1);}
	}
	 tagString+=" [[jd.action]]";
   
   var icreation="";var imodif="";
	 var t=new Date(f(x).CreationTime);
	 icreation=t.convertToYYYYMMDDHHMM();
	 var t=new Date(f(x).LastModificationTime);
	 imodif=t.convertToYYYYMMDDHHMM();
	 
	 var itb=f(x).Body;
	 taskBody=itb;
	 bodyFrom="ol";
	    //TODO get jello notes
	    if (!notEmpty(itb)){taskBody=getJNotesProperty(f(x));bodyFrom="jd";}

	 
	  TWbody+='<div title="'+manualConvertUnicodeToUTF8(f(x).Subject)+'" modifier="'+creator+'" created="'+icreation+'" modified="'+imodif+'" tags="'+manualConvertUnicodeToUTF8(tagString)+'" changecount="1" notes="'+bodyFrom+'" olid="'+f(x).EntryID+'">\n<pre>'+manualConvertUnicodeToUTF8(taskBody)+'</pre></div>\n';
	}

//get tags
var TWtags="";
var icreation=new Date().convertToYYYYMMDDHHMM();var imodif=new Date().convertToYYYYMMDDHHMM();
tagStore.filter("istag",true);
	for (var x=0;x<tagStore.getCount();x++)
	{
	var rc=tagStore.getAt(x);
  var tn=rc.get("tag");
    var isArc=rc.get("archived");
    var isPrj=rc.get("isproject");
    var isPri=rc.get("private");
    var tid=rc.get("id");
    var tno=rc.get("notes");
		var tgtype="[[jd.tag]]";
		if (isPrj){tgtype="[[jd.project]]";}
		if (isArc){tgtype+=" Inactive";}
		if (isPri){tgtype+=" Private";}
		var gtdList="&#60;&#60;gtdActionList&#62;&#62;";
    if (tid==6){tn="jd.next";}
		if (notEmpty(tn))
		{
		TWtags+='<div title="'+manualConvertUnicodeToUTF8(tn)+'" modifier="'+creator+'" created="'+icreation+'" modified="'+imodif+'" tags="'+manualConvertUnicodeToUTF8(tgtype)+'" changecount="1" tagid="'+tid+'">\n<pre>'+gtdList+manualConvertUnicodeToUTF8(tno)+'</pre></div>\n';
		}
	}
tagStore.clearFilter();

//--menus and defaults
var TWother='<div title="DefaultTiddlers" modifier="'+creator+'" created="'+icreation+'" modified="'+imodif+'" tags="jello template" changecount="1">\n<pre>[[Jello.wiki Tag List]]\n[[Next Actions]]</pre></div>\n';
TWother+='<div title="Next Actions" modifier="'+creator+'" created="'+icreation+'" modified="'+imodif+'" changecount="1">\n<pre>&#60;&#60;gtdActionList jd.next&#62;&#62;</pre></div>\n';
TWother+='<div title="Jello Dashboard" modifier="'+creator+'" created="'+icreation+'" modified="'+imodif+'" changecount="1">\n<pre>TiddlyWiki created by Jello Dashboard\n----[[GettingStarted]]\n[[DefaultTiddlers]]</pre></div>\n';
   
    try{
   var twStream = fs.OpenTextFile(destWiki,2,-1,0); }catch(e){alert("Could not create file. Probably folder does not exist yet!");return;}
   var wss=TWheader+"\n"+otherTiddlers+"\n"+TWbody+"\n"+TWtags+"\n"+TWother+"\n"+TWfooter;
   twStream.Write(wss);
   twStream.Close();
   fs=null;
   //update link to the new twiki

var file=fso.getfile(destWiki);
var tsjarfile=fso.getfile(tsaverJAR);
var dfol=file.ParentFolder+"/TiddlySaver.jar";
tsjarfile.Copy(dfol);  
   var dlink=destWiki;
   dlink=dlink.replace(new RegExp("\/","g"),"\//");
jello.lastTiddlyWiki=destWiki;
jese.saveCurrent();
//   Ext.getCmp("twok").hide();



   Ext.getCmp("reslabel2").update("<br><br><br><p align='center'><b>Your TiddlyWiki is Ready!</b><br>&nbsp;<a class=jellolinktop href='"+dlink+"'>Click to open, or Right click here and click Save As</a> to move it to another location</p><br>");

alert("Tiddlywiki file created");  
//--------

}

function passLastTW()
{
var lastf="C:\\Documents and Settings\\nick\\Desktop\\TestingWiki.html";
if (choice=confirm("Sync with last file used? "+lastf)==false){return;}
            var mode=Ext.getCmp("twmode").getValue();
            var folder=getAppPath().replace("jello5.htm","")+"wiki/";
            var file=lastf;
            var tskip=Ext.getCmp("twskip").getValue();
            var sp=Ext.getCmp("skprv").getValue();
            var si=Ext.getCmp("skinc").getValue();
            var onf=Ext.getCmp("olnot").getValue();
            tiddlySync(mode,folder,file,tskip,sp,si,onf);

}

String.prototype.htmlEncode = function()
{
	return this.replace(/&/mg,"&amp;").replace(/</mg,"&lt;").replace(/>/mg,"&gt;").replace(/\"/mg,"&quot;");
};

// Convert "&amp;" to &, "&lt;" to <, "&gt;" to > and "&quot;" to "
String.prototype.htmlDecode = function()
{
	return this.replace(/&amp;/mg,"&").replace(/&lt;/mg,"<").replace(/&gt;/mg,">").replace(/&quot;/mg,"\"");
};


function manualConvertUTF8ToUnicode(utf)
{
return utf.charRefToUnicode();
}


function manualConvertUnicodeToUTF8(s)
{
	var re = /[^\u0000-\u007F]/g ;
	return s.replace(re,function($0) {return "&#" + $0.charCodeAt(0).toString() + ";";});
}


String.zeroPad = function(n,d)
{
	var s = n.toString();
	if(s.length < d)
		s = "000000000000000000000000000".substr(0,d-s.length) + s;
	return s;
};

Date.prototype.convertToYYYYMMDDHHMM = function()
{
	return this.getUTCFullYear() + String.zeroPad(this.getUTCMonth()+1,2) + String.zeroPad(this.getUTCDate(),2) + String.zeroPad(this.getUTCHours(),2) + String.zeroPad(this.getUTCMinutes(),2);
};


Date.convertFromYYYYMMDDHHMM = function(d)
{
	return new Date(Date.UTC(parseInt(d.substr(0,4),10),
			parseInt(d.substr(4,2),10)-1,
			parseInt(d.substr(6,2),10),
			parseInt(d.substr(8,2),10),
			parseInt(d.substr(10,2),10),0,0));
};

function fixTWTags(tg,olit)
{
//convert wiki tag string to jello-outlook format
var td=manualConvertUTF8ToUnicode(tg);
var ns=td.replace("jd.action","");
ns=Trim(ns);
var o=ns.split(" ");
var catString="";
var sflag=false;
    var ff="[[";
    var gg="]]";
  for (var x=0;x<o.length;x++)
  {
    var ss=o[x];

    var s=ss.search(ff.escapeRegExp());
    if (s>-1)
    {sflag=true;}
    
      var d=ss.search(gg.escapeRegExp());
      if (d>-1){sflag=false;}

      if (sflag){catString+=o[x]+" ";}else{catString+=o[x]+";";}    
  }
  
  catString=catString.replace(new RegExp(ff.escapeRegExp(),"g"),"");
  catString=catString.replace(new RegExp(gg.escapeRegExp(),"g"),"");
 	var nextTag="!Next";
  try{var r=getTagByID("6");nextTag=r.get("tag");}catch(e){}
  catString=catString.replace("jd.action;","");
  catString=catString.replace("jd.next;",nextTag+";");
  
  //check for jd.done tag to complete the action
  s=catString.search("jd.done");
  var isDone=false;
  if (s>-1){isDone=true;}
  if (isDone && olit.Status!=2){olit.Status=2;olit.Save();}
  catString=catString.replace("jd.done;",nextTag+";");
  
  if (catString==";"){catString="";}

//todo- check category string for newly added tags and add them into jello
if (!notEmpty(catString)){return catString;}
var cts=catString.split(";");
if (cts.length==0){return catString;}

  for (var x=0;x<cts.length;x++)
  {
   var cat=Trim(cts[x]);
   if (cat!=" " && notEmpty(cat)){
	var r=getTag(cat);
	if (r==null){createTag(cat,true,0,false,false,false);}
   }
  }
    
  return catString;
}

// Escape any special RegExp characters with that character preceded by a backslash
String.prototype.escapeRegExp = function()
{
	var s = "\\^$*+?()=!|,{}[].";
	var c = this;
	for(var t=0; t<s.length; t++)
		c = c.replace(new RegExp("\\" + s.substr(t,1),"g"),"\\" + s.substr(t,1));
	return c;
};

String.prototype.charRefToUnicode = function()
{
return this.replace(
/&#(([0-9]{1,7})|(x[0-9a-f]{1,6}));?/gi,
function(match, p1, p2, p3, offset, s)
{
return String.fromCharCode(p2 || ("0" + p3));
});
}

function getJNotesProperty(it){
 // duplicate function so you dont have to load actionlist.js
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
