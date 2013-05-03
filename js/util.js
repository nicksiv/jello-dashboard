//util.js

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



function DisplayCurrentDate() {
//**j5
var d=new Date();
return DisplayDate(d);
}

function Trim(S) {
//**j5
var p1;
var p2;
var l;
//if (typeof(l)=="undefined"){return "";}
try{
l = S.length;
if (l==0) {return '';}
for (p1=0;p1<l;p1++) {if (S.charAt(p1) != ' ') {break;}}
for (p2=l-1;p2>0;p2--) {if (S.charAt(p2) != ' '){break;}}
if (p2<p1) {return '';}
return S.substring(p1,p2+1);
}catch(e){return S;}
}

function DisplayDate(t,showOLNullDate) {
//**j5
// if ShowOLNullDate is passed, even return a string for
// the OL null date, 1/1/4501
var doNull = false;
if( DisplayDate.arguments.length == 2)
	if (showOLNullDate!=null){doNull = showOLNullDate;}
if(!notEmpty(t)){return txtUndefined;}
try{
if (t.getYear() == 4501 && t.getMonth() == 0 && t.getDate() == 1 && !doNull) {return "";}
}catch(e){}
// true says use long format
return t.format(jelloDateFormatString(true));
//if (jello.dateFormat == 0 || jello.dateFormat == "0") {return t.getDate() + jello.dateSeperator + (t.getMonth()+1) + jello.dateSeperator + t.getYear() + "" ;}
//return (t.getMonth()+1) + jello.dateSeperator + t.getDate() + jello.dateSeperator + t.getYear() + "" ;

}
function DisplayAppDate(t,dayOnly) {
if (t==null || t==""){return "";}
if (t.getYear() == 4501 && t.getMonth() == 0 && t.getDate() == 1) {return "";}
var tod=new Date();
var s=dayDif(t,tod,true);
ss="<span class=dayfuture>";
if (s<1){ss="<span class=daypdue>";}
if (s==0){ss="<span class=daytoday>";}
if (s==1){ss="<span class=dayfuture>";}
if (jello.dueDisplayDaysLeft==false){
if (jello.dueShowYear==true){var yr=jello.dateSeperator+t.getFullYear();}else{var yr="";}

var	dow = dayOfWeek(t.getDay(),1);

if( typeof(dayOnly) != "undefined" && dayOnly != null && dayOnly == true)
	return ss + dow + "<span>";

if (jello.dateFormat == "0" || jello.dateFormat == 0) {return ss + dow+ " " + t.getDate() + jello.dateSeperator + (t.getMonth()+1)  + yr+"</span>" ;}
if (jello.dateFormat == "2" || jello.dateFormat == 2) {return ss + dow+ " " + yr + (t.getMonth()+1) + jello.dateSeperator + t.getDate() +"</span>" ;}
return  ss + dow+ " " + (t.getMonth()+1) + jello.dateSeperator + t.getDate() + jello.dateSeperator + yr+ "</span>" ;


if (jello.dateFormat == "0" || jello.dateFormat == 0) {return ss + dayOfWeek(t.getDay(),1)+ " " + t.getDate() + jello.dateSeperator + (t.getMonth()+1) +   yr+"</span>" ;}
if (jello.dateFormat == "2" || jello.dateFormat == 2) {return ss + yr+ " " + t.getDate() + jello.dateSeperator + (t.getMonth()+1) + jello.dateSeperator + dayOfWeek(t.getDay(),1) +"</span>" ;}
return  ss + dayOfWeek(t.getDay(),1)+ " " + (t.getMonth()+1) + jello.dateSeperator + t.getDate() + jello.dateSeperator + yr+ "</span>" ;
}
else
{
//return days left
var tod=new Date();
return dayDif(t,tod);
}
}

function DisplayTime(t) {
//**j5
if (t.getYear() == 4501 && t.getMonth() == 0 && t.getDate() == 1) {return "";}
var m = ((t.getMinutes() < 10) ? '0' : '') + t.getMinutes();
if (jello.timeFormat == 0) {return t.getHours() + ":" + m ;}
if (t.getHours() > 12) {return (t.getHours() - 12) + ":" + m + ' PM';}
if (t.getHours() == 12 && t.getMinutes() >=0) {return (t.getHours()) + ":" + m + ' PM';}
return t.getHours() + ":" + m + ' AM' ;
}


function secondDif(today,num)
{
/*---------------------------------------calc seconds till Date*/
var t=new Date(today.getYear(),today.getMonth(), today.getDate() + num);
/*var xx= (Math.ceil((today.getHours()-t.getHours())));*/
var diff  = new Date();
diff.setTime(Math.abs(today.getTime() - t.getTime()));
var timediff = diff.getTime();
return (timediff/1000);
}

function dayDifSimple(day1,day2)
{
//calc difference in days between two dates
var dif=0;
try{var one_day=1000*60*60*24;
var dif=Math.ceil((day1.getTime()-day2.getTime())/(one_day));}catch(e){}
return dif;
}

function daysInMonth(month,year) {
var dd = new Date(year, month, 0);
return dd.getDate();
}

function dayDif(nextDate,t,no)
{
/*Calculate day difference*/
var r=0;
nextDate.setHours(0,0,0,0);
t.setHours(0,0,0,0);
var ret="<span class=textgreen>("+txtAllDay+")</span>";
t=new Date();
var one_day=1000*60*60*24;
var xx= (Math.ceil((nextDate.getTime()-t.getTime())/(one_day)));
//xx--;
var d1=parseInt(t.getDate());
var d2=parseInt(nextDate.getDate());
var m1=(t.getMonth()+1);
var m2=(nextDate.getMonth()+1);
var y1=(t.getYear());
var y2=(nextDate.getYear());
//if ((d1-d2)==0 && m1==m2 && y1==y2){xx=0;}
if (xx==0) {ret="<span class=today>"+txtToday+"</span>&nbsp;&lt;&lt;";r=0;}
if (xx==1) {ret="<span class=info2>"+txtTomorrow+"</span>";r=1;}
if (xx>1) {ret="<span class=textNormal>" + (xx)+ " "+txtDays+"</span>";r=xx;}
if (xx<0){ret="<span class=counter>"+(xx-(xx*2))+" "+txtdaysAgo+"</span>";r=-1;}
if (no==true){return r;}else{return ret;}
}

function trimLow(ss)
{  //**j5
try{
var e=Trim(ss);
return e.toLowerCase();
}catch(e){return ss;}
}

function dayOfWeek(d,small)
{
//returns day of the week
try{var a=txtDayList.split(",");
var b=a[d];
if (small==1){b=b.substr(0,3);}
return b;}catch(e){return "";}
}

function monthName(d)
{
//returns month's name
var a=txtMonthList.split(",");
//d++;
if (d==12){d=0;}
var b=a[d];
return b;
}

function htmlEntities(str) {
    var i,output="",len,charc="";
    len = str.length;
    if (len>0){
    for(i=0;i<len;i++){
        charc = str.charCodeAt(i);
        if( (charc>47 && charc<58)||(charc>62 && charc<127) ){
            output += str.substr(i,1);
        }else{
            output += "&#" + str.charCodeAt(i) + ";";
        }
    }}
    return output;
}

function notEmpty(s)
{
//check string for empty or null or space
var ret=true;
if (s==null || s=="" || s==" " || typeof(s)=="undefined" || s=="NaN"){ret=false;}
return ret;
}

function isDate(str)
{
//check if string is a date
            var isDate=false;
            var dt=new Date(str);
            if (dt!="NaN"){isDate=true;}
            return isDate;
}


function help()
{
//display the help file
alert("Fix this");
}

function arrayExists(vl,array)
{
//check if value exists in array
var exs=false;
	if (array.length>0)
	{
		for (var x=0;x<array.length;x++)
		{
			if (Trim(array[x])==Trim(vl)){exs=true;break;}
		}
	}
return exs;
}

function getAppPath()
{//return application path
   var uri="";
   if (conStatus!="Outlook ActiveX")
   {uri=ol.ActiveExplorer.CurrentFolder.WebViewURL.replace("jello5.htm","");}
   else
   {uri=location.href;uri=uri.replace("jello5.hta","");
   uri=uri.replace(new RegExp("%20","gi")," ");
   uri=uri.replace("file:///","");
   }
   return uri;
}

function valExistsInArray(vl,theArray)
{
	var found=false;
	if (theArray.length>0)
	{
	for (var x=0;x<theArray.length;x++)
		{
		if (theArray[x]==vl){found=true;break;}
		}
	}
return found;
}

function rightClickItemMenu(e,row,g)
{
	e.stopEvent();
	var store = g.getStore();
	var itms = getCheckedItems(null);
	var rec = store.getAt(row);
	var iclass = rec.get("iclass");
	var id = rec.get("entryID");
	// is the item right clicked a selected item
	var inlist=false;
	// walk though the selected and look for maching id
	for( var i=0; i < itms.length; i++)
	{
		if( id != itms[i].get("entryID"))
			continue;
		inlist = true;
		break;
	}
	var cmnu = null;


	if( g.getId() == "grid" || iclass == 48){
		cmnu = actionContextMenu(rec,id,inlist);
	}else if (iclass==43 || iclass==45){
		try{it = NSpace.GetItemFromID(id);
		cmnu = mailContextMenu(rec,id, it, inlist);}catch(e){}
	}else if( iclass == 26 || iclass == 53){
		try{it = NSpace.GetItemFromID(id);
		cmnu = ticklerContextMenu(rec,id, it, inlist);}catch(e){}
	}
	if( cmnu != null) {cmnu.showAt(e.getXY());}

}

function toHtmlLines(s){
	s = s.replace(/\r/g,"");
	return s.replace(/\n/g,"<br>\n");
}
function jelloDateFormatString(longyear)
{
var year = "y";
if(jelloDateFormatString.arguments.length != 0)
	year = "Y";
var dts=jello.dateSeperator;
var fmt="m"+dts+"d"+dts+ year;
if(jello.dateFormat == 0 || jello.dateFormat == "0"){fmt="d"+dts+"m"+dts+year;}
if(jello.dateFormat == 2 || jello.dateFormat == "2"){fmt=year+dts+"m"+dts+"d";}
return fmt;
}

function createDonationAction()
{
try{
 var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
 var r=getTagByID(6);
 var dt=iF.Add();
 dt.Subject="Readme:Donate to Jello Dashboard";
 var msg="Thank you for trying Jello Dashboard! If you're enjoying using this application please Donate to support its development. This is a one time only notification. The application will not bother you again. Huge thanks if you've already donated! Thank you, the developer.";
 dt.Body=msg;
 dt.Importance=0;
 if (r!=null){dt.itemProperties.item(catProperty).Value=r.get("tag");}
 setJNotesProperty(dt,msg);
 dt.Save();
 jello.firstRun=0;jese.saveCurrent();
}catch(e){}
}

function dynamicLoad(scrp,fun)
{

try{
eval(fun);
    }
    catch(e)
    {
   var head= document.getElementsByTagName('head')[0];
   var script= document.createElement('script');
   script.type= 'text/javascript';
   script.onreadystatechange= function () {
      if (this.readyState == 'complete') eval(fun);
   };
   script.onload=fun;
   script.src= 'js/'+scrp;
   head.appendChild(script);

   }

}

function getPercentIcons(num,count,aall)
{
//returns a series of icons showing visually a percentage

if (aall>0 && count==0){return "<img src=img\\check.gif title='"+txtCompleted+"'>";}
if (num==0 && count==0){return "<img title='No actions' src=img\\clock.gif style='opacity:0.4;filter:alpha(opacity=40);'>";}
var n=parseInt(num/10);
if (n==10){return "<img src=img\\check.gif>";}
var ret="";
	for (var x=0;x<n;x++)
	{ret+="<img src=img\\action_go.gif style='width:10px;opacity:1;filter:alpha(opacity=100);' title='"+count+" "+txtCbActions+" "+parseInt(num)+"%'>";}

	if (n<10)
	{
		for (var x=n;x<10;x++)
		{ret+="<img src=img\\action_go.gif style='width:10px;opacity:0.5;filter:alpha(opacity=50);' title='"+count+" "+txtCbActions+" "+parseInt(num)+"%'>";}

	}

	return ret;
}

var Browser = {
  Version: function() {
    var version = 999;
    if (navigator.appVersion.indexOf("MSIE") != -1)
       version = parseFloat(navigator.appVersion.split("MSIE")[1]);
    return version;
  }
};

function hourDif(day1,day2)
{
//calc difference in days between two dates
var dif=0;
try{var one_hour=1000*60*60;
var dif=Math.ceil((day1.getTime()-day2.getTime())/(one_hour));}catch(e){}
return dif;
}

//----justausrcode
function checkForSmartDate(df, value)
{
var sword;
var ret = false;
if( typeof(df) != "undefined" && df != null)
        if (Ext.isDate(df.getValue()))
        {//normal datevalue
        return;
        }
        else
        {//other
        sword=df.getRawValue();
        }
else{
        sword = value;
        ret = true;
}

var nnd=null;
sword = Trim(sword);
var ims = sword.split(" ");
var x = new LineObject();
if( x.length !=0)
{
	for(var i=0; i < ims.length; i++){
		var el=Trim(ims[i]);
		if (!notEmpty(el))
				continue;
		x.nextWord(el);
	}
	if( x.startDate == "")
		x.noiseToDesc();
	nnd = x.startDate;


}

	if (notEmpty(nnd)){
		// 2 or 4 digit year
        if( ret)
            return nnd;
        else
		df.setValue(nnd);
	}else{if (!isDate(sword)){
        if( ret)
            return null;
        else
            df.setValue(null);}
    }

}

function checkForSmartDatebase(df)
{
var sword=df.getRawValue();
var nd;var nnd=null;
var dts=jello.dateSeperator;
var fmt="n"+dts+"j"+dts+"y";
if(jello.dateFormat == 0 || jello.dateFormat == "0"){fmt="j"+dts+"n"+dts+"y";}
if(jello.dateFormat == 2 || jello.dateFormat == "2"){fmt="Y"+dts+"m"+dts+"d";}
if (sword=="today" || sword=="now" || sword=="tod"){nd=new Date();nnd=nd.format(fmt);}
if (sword=="tomorrow" || sword=="tom"){nd=new Date();nd.setDate(nd.getDate()+1);nnd=nd.format(fmt);}
if (sword=="next week" || sword=="nw"){nd=new Date();nd.setDate(nd.getDate()+7);nnd=nd.format(fmt);}
if (sword=="next month" || sword=="nm"){nd=new Date();nd.setDate(nd.getDate()+30);nnd=nd.format(fmt);}
if (sword=="next year" || sword=="ny"){nd=new Date();nd.setDate(nd.getDate()+365);nnd=nd.format(fmt);}

	//in x days,weeks,months
	if (sword.substr(0,3)=="in ")
	{
		var o=sword.split(" ");
		try
		{var num=parseInt(o[1]);
			if (typeof(num)=="number")
			{
				var itv=1;
				if (trimLow(o[2].substr(0,1))=="m"){itv=30;}
				if (trimLow(o[2].substr(0,1))=="w"){itv=7;}
				if (trimLow(o[2].substr(0,1))=="y"){itv=365;}
				var pls=num*itv;
				nd=new Date();nd.setDate(nd.getDate()+pls);nnd=nd.format(fmt);
			}
		}
		catch(e){}
	}

//day names
var daylist=txtDayList.split(",");
var theday=null;
	for (var x=0;x<daylist.length;x++)
	{
		if (trimLow(sword.substr(0,3))==trimLow(daylist[x].substr(0,3)))
		{theday=x;break;}
	}
if (theday!=null)
{nd=new Date();
var tdd=nd.getDay();
if (tdd<theday){nd.setDate(nd.getDate()+(theday-tdd));nnd=nd.format(fmt);}
if (tdd==theday){nd.setDate(nd.getDate());nnd=nd.format(fmt);}
if (tdd>theday){nd.setDate(nd.getDate()+((theday-tdd)+7));nnd=nd.format(fmt);}
}

if (notEmpty(nnd)){df.setValue(nnd);}else{if (!isDate(sword)){df.setValue(null);}}
}
//------------------

//----drubqarcode
function drscheckForSmartDate(df)
{
var sword=df.getRawValue();
var nd;var nnd=null;
var dts=jello.dateSeperator;
var fmt="n"+dts+"j"+dts+"y";
if(jello.dateFormat == 0 || jello.dateFormat == "0"){fmt="j"+dts+"n"+dts+"y";}
if(jello.dateFormat == 2 || jello.dateFormat == "2"){fmt="Y"+dts+"m"+dts+"d";}
if (sword=="today" || sword=="now" || sword=="tod"){nd=new Date();nnd=nd.format(fmt);}
if (sword=="tomorrow" || sword=="tom"){nd=new Date();nd.setDate(nd.getDate()+1);nnd=nd.format(fmt);}
if (sword=="next week" || sword=="nw"){nd=new Date();nd.setDate(nd.getDate()+7);nnd=nd.format(fmt);}
if (sword=="next month" || sword=="nm"){nd=new Date();nd.setDate(nd.getDate()+30);nnd=nd.format(fmt);}
if (sword=="next year" || sword=="ny"){nd=new Date();nd.setDate(nd.getDate()+365);nnd=nd.format(fmt);}

	//in x days,weeks,months
	if (sword.substr(0,3)=="in ")
	{
		var o=sword.split(" ");
		try
		{var num=parseInt(o[1]);
			if (typeof(num)=="number")
			{
				var itv=1;
				if (trimLow(o[2].substr(0,1))=="m"){itv=30;}
				if (trimLow(o[2].substr(0,1))=="w"){itv=7;}
				if (trimLow(o[2].substr(0,1))=="y"){itv=365;}
				var pls=num*itv;
				nd=new Date();nd.setDate(nd.getDate()+pls);nnd=nd.format(fmt);
			}
		}
		catch(e){}
	}

//day names
var daylist=txtDayList.split(",");
var theday=null;
	for (var x=0;x<daylist.length;x++)
	{
		if (trimLow(sword.substr(0,3))==trimLow(daylist[x].substr(0,3)))
		{theday=x;break;}
	}
if (theday!=null)
{nd=new Date();
var tdd=nd.getDay();
if (tdd<theday){nd.setDate(nd.getDate()+(theday-tdd));nnd=nd.format(fmt);}
if (tdd==theday){nd.setDate(nd.getDate());nnd=nd.format(fmt);}
if (tdd>theday){nd.setDate(nd.getDate()+((theday-tdd)+7));nnd=nd.format(fmt);}
}

if (notEmpty(nnd)){df.setValue(nnd);}else{if (!isDate(sword)){df.setValue(null);}}
}

