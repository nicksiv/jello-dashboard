// Review module

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


var tRvReload={
  id:'refresh',
  handler: function(e, target, panel){
  Ext.getCmp("reviewacpanel").expand(true);
  updateReviewPanel(); 
  }        
  };

  var tRvNext={
  id:'right',
  qtip:txtNext + ' Ctrl+Page Down',
  handler: function(e, target, panel){
  Ext.getCmp("reviewacpanel").expand(true);
  reviewGo(1); 
  }        
  };

  var tRvPrev={
  id:'left',
  qtip:txtPrevious + ' Ctrl+Page Up',
  handler: function(e, target, panel){
  Ext.getCmp("reviewacpanel").expand(true);
  reviewGo(-1); 
  }        
  };
  
  var tRvEmpty={
  id:'revempty',
  qtip:txtShowEmptyTags,
  handler: function(e, target, panel){
  showZeroTags();
  }        
  };

  var tRvPath={
  id:'revpaths',
  qtip:txtShowPaths,
  handler: function(e, target, panel){
  if (jello.reviewPanelPaths==1 || jello.reviewPanelPaths=="1"){jello.reviewPanelPaths=0;}else{jello.reviewPanelPaths=1;}
  jese.saveCurrent();
  Ext.getCmp("reviewacpanel").expand(true);
  updateReviewPanel(); 
  }        
  };
    
var reviewStore;
var currentTag=0;
var showZero=jello.reviewShowEmptyTags;
var panelItems=new Array();
var panelPJItems=new Array();

function updateReviewPanel(tgName)
{
//update review panel of sidebar
if (typeof(tgName)=="number"){tgName=getTagName(tgName);}
if (showZero){btti=txtHideEmptyTags;}
var showPaths=false;
if (jello.reviewPanelPaths==1 || jello.reviewPanelPaths=="1"){showPaths=true;}
panelItems=new Array();
currentTag=-1;
try{
var tagName=tgName.replace(new RegExp("~","g")," ");
var ctid=getTagID(tagName);
 }catch(e){var ctid=0;}
 
status="Creating tag list";
reviewStore=tagStore;
var ret="<table cellspacing=0 cellpadding=0>";
var tgCount=0;

    var pathTag=function(r){
    var tname=r.get("tag");
    var par=r.get("parent");
    var caption=Trim(TagFullPath(par,tname,false));
    r.beginEdit();r.set("path",caption);r.endEdit();
    };

    var renderTag = function(r){
      var tname=r.get("tag");
      var tid=r.get("id");
      var tpv=r.get("private");
      var istg=r.get("istag");
      var cls=r.get("closed");
      if (!istg || cls){return;}
      var tarc=r.get("archived");
      if (tpv){return;}
      var tgname=tname.replace(new RegExp(" ","g"),"~");
      var caption="";
      if (showPaths){caption=r.get("path");}else{caption=tname;}
      
      if (caption==getActionTagName()){caption=txtNextActions;}
	caption=Ext.util.Format.htmlEncode(caption);
      var tagLink="<a class=jellolinkreview onclick=showContext("+tid+",1);updateReviewPanel("+tid+");>"+caption+"</a>";
if (jello.reviewpopup==1 || jello.reviewpopup=="1")
{tagLink += "&nbsp;&nbsp;<a title='Show in popup' class=jellolinktop onclick=tasksWithTag(null,'"+tgname+"'); ><font size='-1' face=webdings>1</font></a>";}
      var tCount=showContext(tname,0);
      if (tCount>0){var tCounter=tCount;}else{if (!showZero){return;}else{var tCounter="&nbsp;";}}
	 panelItems.push(tid);
      var fmt="";if (tid==ctid){currentTag=(panelItems.length-1);fmt="class=tagprj";}
      ret+="<tr><td width=5% height=22 style='border-bottom:1px dotted gray;' valign=middle>"+actionListTitleIcon(tname)+"</td>";
      ret+="<td style='border-bottom:1px dotted gray;' valign=middle width=90%><span "+fmt+">"+tagLink+"</span></td>";
      ret+="<td style='border-bottom:1px dotted gray;' width=5%><span class=fkey>"+tCounter+"</span></td></tr>";
            tgCount++;
    };

status=txtStatusGetting+"...";

if (showPaths){reviewStore.each(pathTag);}

if (showPaths){reviewStore.sort("path","ASC");}else{reviewStore.sort("tag","ASC");}
reviewStore.each(renderTag);
ret+="</table>";
status=txtReady;
Ext.getCmp("reviewacpanel").update(ret);
Ext.getCmp("reviewacpanel").setTitle(txtReview2+" ("+tgCount+")");
}

function pReview()
{
//*********
Ext.getCmp("accord").setActiveTab("reviewpanelparent");    
updateReviewPanel();
}

function showZeroTags()
{//update the panel to show/not show empty tags

            showZero=!showZero;
            jello.reviewShowEmptyTags=showZero;
            jese.saveCurrent();
            updateReviewPanel();
}


function reviewGo(num)
{
//move reviewd tag
//currentTag
status="Fetching next tag...";
Ext.getCmp("accord").setActiveTab("reviewpanelparent");

var tCount=1;

if (num>0)
	{
	currentTag++;
	if ((panelItems.length-1)<currentTag){currentTag=0;}
	var goRec=getTagByID(panelItems[currentTag]);
	}
	
if (num<0)
	{
	currentTag--;
	if (currentTag<0){currentTag=panelItems.length-1;}
	var goRec=getTagByID(panelItems[currentTag]);
	}	

	
	var ntg=goRec.get("tag");
	var tpv=goRec.get("private");
	showContext(ntg,1);updateReviewPanel(ntg);
status="";	
}
