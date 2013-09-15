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
//    2008-2013 N.Sivridis http://jello-dashboard.net


//variable with last filename : lastTodotxt, lastsync: todotxtLastSync

var todoDefaultName="todo.txt";
var SNextAction=getActionTagName();
var txtSpaceChar="_";
var txtFile;
var newTxtFile;
var writeFile;

var txtLineSeparator=10;
//windows linefeed  -1, 13
// linux 10 (works)

var stream;
var txtDateDelimiter="t";

var txtStore=new Ext.data.JsonStore({
        fields: [
	   {name: 'lineid',type:'int'},
           {name: 'priority'},
           {name: 'done',type:'boolean'},
           {name: 'tags', type:'string'},
           {name: 'created', type:'date'},
           {name: 'donedate', type:'date'},
           {name: 'task'},
           {name: 'extradata'},
           {name: 'oldue', type:'date'},
           {name: 'olid'}
       ]
    });


 var txtRecord = Ext.data.Record.create([
	   {name: 'lineid',type:'int'},
           {name: 'priority'},
           {name: 'done',type:'boolean'},
           {name: 'tags', type:'string'},
           {name: 'created', type:'date'},
           {name: 'donedate', type:'date'},
           {name: 'task'},
           {name: 'extradata'},
           {name: 'oldue', type:'date'},
           {name: 'olid'},
]);


function todoTxtSync()
{
//Todotxt sync main function


    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");

  
var simple = new Ext.FormPanel({
        frame:true,
    	title:'Todo.txt file Export',
        bodyStyle:'padding:5px 5px 0 30px',
        floating:false,
        labelWidth:175,
        id:'txtform',
        buttonAlign:'center',
        defaultType: 'textfield',

        items: [
	       

	    {
                fieldLabel: 'Full path of Todo.txt file',
                inputType:'text',
                id: 'ttwsync',
                hidden:false,
                tabIndex:2,
	//	listeners:{blur:function(a){getTWFileInfo(false);}},
		width:420,
                allowBlank:true
            },
	    	new Ext.form.Label({
		html:'',
		width:450,
	    	id:'reslabel',
		cls:'formfilelist',
		fieldLabel:' ',labelSeparator:'',
		height:50
		}),
             /*
            {
                fieldLabel: 'Skip tags (comma separated)',
                id: 'twskip',
                visible:true,
                value:'',
                tabIndex:3,
		width:420,
                allowBlank:true
            },*/
        new Ext.form.Label({
				html:'',
				width:650,
	    			id:'reslabel2',
				cls:'formfilelist',
            			height:50
				})
	    ,new Ext.Button({text:'<b>Export Todo.txt File</b>',id:'twgo',fieldLabel:' ',autoWidth:false,width:180,height:30,labelSeparator:'',listeners:{click:function(p){todoSync();}}})

        ]

    });


simple.render(mp.getEl());

    setTimeout(function(){
//	getTWFileInfo(true);
    }, 250);

if (init)
{
    if (!notEmpty(jello.lastTodotxt) || (jello.lastTodotxt=="undefined"))
	{
	    var jepath="";
	    var fso = new ActiveXObject("WScript.Shell");
	    jepath = fso.SpecialFolders("MyDocuments")+"\\"+todoDefaultName;
	    jepath = jepath.replace(/\\/g, "\\");
	}
    else
	{var jepath=jello.lastTodotxt;}

    Ext.getCmp("ttwsync").setValue(jepath);
}



}



function todoSync()
{
txtStore.removeAll();

var fso = new ActiveXObject("Scripting.FileSystemObject");
var fname=Ext.getCmp("ttwsync").getValue();
try{
    txtfile=fso.GetFile(fname);}catch(e){if (!confirm("File "+fname+" does not exist \n Do you want to create it?")){return;}else
					  {  
					     
					      try{fso.CreateFolder(fso.GetFileFolder(fname));}catch(g){}
					     try{
					      fso.CreateTextFile(fname,-1,-1);
					      txtfile=fso.GetFile(fname); }catch(f){alert("Something went wrong. File could not be created!\n"+f.description);return;}
					  }}

var txtLines=new Array();

//open the todo.txt file and get lines
/*var ts = txtfile.OpenAsTextStream(1,-2);
var s="";

if (!ts.AtEndOfStream){
    do
	{
	    s=ts.ReadLine();
	    txtLines.push(s);
	} while (!ts.AtEndOfStream)	
}
ts.close();*/

stream = new ActiveXObject("adodb.stream");
    stream.Open;
    stream.Type=2;
    stream.CharSet="UTF-8";  
    stream.LineSeparator=txtLineSeparator;
    stream.LoadFromFile(txtfile.Path);
    var s="";
    if(!stream.EOS)
    {
      do
        {
          s=stream.ReadText(-2);
          txtLines.push(s);
        } while (!stream.EOS)
    }


    //01. read each line, handle and add to a table of values
    // Export only for the time being
    //for (var x=0;x<txtLines.length;x++)
   //{handleTxtLine(txtLines[x],x);}
    
    //02. read all outlook jello items, compare and add to the table
    var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
    var dasl="http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2"
    var its=iF.Restrict("@SQL="+dasl);
    var counter=its.Count;
    	for (var y=1;y<=counter;y++)
	    {
		handleOLItemToTxt(its(y));
	    }
    
    //03. loop through txtStore table, add items with no olid to Outlook and update text file
    //txtStore.sort("lineid","ASC");
    stream = new ActiveXObject("adodb.stream");

    var dfol=txtfile.ParentFolder+"/todo-jello.bak";
    txtfile.Copy(dfol);  

    stream.Open;
    stream.Type=2;
    stream.CharSet="UTF-8";  
    stream.LineSeparator=txtLineSeparator;
   // stream.LoadFromFile(txtfile.Path);
    
    txtStore.each(loopTxtItems);
    stream.SaveToFile (txtfile.Path,2);
    stream.close();

	  
//End and info
    jello.lastTodotxt=txtfile.Path;
    jese.saveCurrent();
    
    Ext.getCmp("reslabel2").update("Export successful<hr><h2><b>"+txtStore.getCount()+" Items exported</b></h2>");
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


function isDate(str)
{
//check if string is a date
            var isDate=false;
//            var dt=new Date(str);
//            if (dt!="NaN"){isDate=true;}
    		var dt=new Date();
    		dt=Date.parseDate(str,"Y-m-d")
    		if (dt!=undefined){isDate=true;}
            return isDate;
}

var DateDiff = {

    inDays: function(d1, d2) {
        var t2 = d2.getTime();
        var t1 = d1.getTime();

        return parseInt((t2-t1)/(24*3600*1000));
    }
}

function handleTxtLine(tLine,num)
{
//add fields of txt line to table

if (tLine=="" || tLine==" " || tLine=="  " || tLine.length<5){return;}
//split the todotxt line in spaces
var items=tLine.split(" ");

var fpri="";
var ftask="";
var fdone=false;
var fdonedate="";
var fcreated="";
var ftags="";
var ftask="";
var fdue="";

    //priority item 
    if (items[0].substr(0,1)=="(")
	{fpri=items[0].substr(1,1);items[0]="";
    if (isDate(items[1])){fcreated=items[1];items[1]="";}

	}
    	
    //done item
    if (items[0]=="x")
    {fdone=true;fdonedate=items[1];
     if (isDate(items[2])){fcreated=items[2];items[2]="";}
     items[0]="";items[1]="";items[2]="";}

    //simple task with creation date
    if (isDate(items[0])){fcreated=items[0];items[0]="";}
    
    //tags
    for (var x=0;x<items.length;x++)
	{
	    if (items[x].substr(0,1)=="@" || items[x].substr(0,1)=="+")
		{ftags+=items[x].toLowerCase()+" ";items[x]="";}
	
		//due date previously entered from outlook!	
		if (items[x].substr(0,txtDateDelimiter.length+1)==txtDateDelimiter+":")
		{fdue=items[x].substr(txtDateDelimiter.length,10); items[x]="";}
	}

	


    ftask=items.join(" ");
	
var tr=new txtRecord({
priority:fpri,
done:fdone,
donedate:Date.parseDate(fdonedate,"Y-m-d"),
created:Date.parseDate(fcreated,"Y-m-d"),
tags:ftags,
task:Trim(ftask),
lineid:num,
olid:isInOutlook(Trim(ftask))
});
txtStore.add(tr);


}

function isInOutlook(taskName)
{
//Search for task by name in Outlook
//return itemID if found
var isThere=0;
taskName=taskName.replace(new RegExp("'","g"),"''");                                                                                   
var dasl="(http://schemas.microsoft.com/mapi/proptag/0x0037001f = '" + taskName + "') AND http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2";
var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
  try{
  var its=iF.Restrict("@SQL="+dasl);
  counter=its.Count;
  if (counter>0) {isThere=its(1).EntryID;}
  }catch(e){alert("Cannot handle task string"+taskName);}

return isThere;
}

function JelloTagsToTxt(cats)
{
var tTags="";
var nextFlag=false;
cats=cats.toLowerCase();
cats=cats.replace(new RegExp(",","g"),";");
var ee=cats.split(";");
			for (var x=0;x<ee.length;x++)
			{
			    	var o=Trim(ee[x]);
				if (trimLow(o)!=trimLow(SNextAction))
			    	{
				    var r=getTag(o);var tagPre="@";
				    var isp=r.get("isproject");
				    if (isp){tagPre="+"}
				    o=o.replace(new RegExp(" ","g"),txtSpaceChar);
				    if (tagPre!=o.substr(0,1)){o=tagPre+o;}
				    tTags=tTags+o+" ";
				}else{nextFlag=true;}
			}
if (nextFlag){tTags="{{{*}}}"+tTags;}
return tTags;
}

function JelloNextActionToTxt(cats)
{
//return true if there is a !Next tag in category string
isNext=false;
cats=cats.replace(new RegExp(",","g"),";");
var ee=cats.split(";");
			for (var x=0;x<ee.length;x++)
			{
			    	var o=Trim(ee[x]);
				if (trimLow(o)==trimLow(SNextAction))
			    	{
				   isNext=true;return isNext; 
				}
			}

return isNext;
}

function handleOLItemToTxt(olitem)
{
//handle an outlook item

var fpri="";
var ftask="";
var fdone=false;
var fdonedate="";
var fcreated="";
var ftags="";
var ftask="";
var hasNext=false;

    try{
    var ix=txtStore.find("olid",new RegExp(olitem.EntryID));
        }catch(e){alert("Error in "+olitem.Subject);}
    var r=txtStore.getAt(ix);
    	if (r==undefined || r==null)
	 {
	     //outlook item not found in txt. Add it
	     fpri="";fcreated=olitem.CreationTime;
	     ftags=JelloTagsToTxt(olitem.itemProperties.item(catProperty).Value);
       hasNext=JelloNextActionToTxt(olitem.itemProperties.item(catProperty).Value);
	     
       ftask=Trim(olitem.Subject);
	     
       var opri=getJPriorityProperty(olitem.EntryID);
	     //** high Priority Outlook items get pri B in todotxt 
       if (opri=="1"){fpri="B";}
	     //if (opri=="2"){fpri="C";}
	     //if (opri=="3"){fpri="D";}
	     
       //Outlook next actions get pri A in todotxt
       if (hasNext){fpri="A";}
       
       if (ftags.substr(0,7)=="{{{*}}}"){fpri="A";ftags=ftags.substr(7,ftags.length);}
	     //if there is a due date add it after tags and update priority
	     var fdue=new Date(olitem.DueDate);
	     if (fdue.getFullYear()!=4501)
	     	{ftags+=txtDateDelimiter+":"+fdue.getFullYear()+"-"+fdue.getMonth()+"-"+fdue.getDate();
    		 var tdays=DateDiff.inDays(new Date(),fdue);
		 // This should be optional?
     if (tdays<3){opri="B";}
		 if (tdays<8 && tdays>2){opri="C";}
		 if (tdays<12 && tdays>7){opri="D";}
		 if (tdays<21 && tdays>13){opri="E";}
		}

	     var tr=new txtRecord({
		 priority:fpri,
		 created:fcreated,
		 tags:ftags,
		 task:Trim(ftask),
		 olid:olitem.EntryID
	     	});
	     txtStore.add(tr);

	 }
    	else
	 {
	     //outlook item found in txt. Need to spot differences?
	 }
}


var loopTxtItems= function(r,tStream){
//handle each txt item
var tl="";
    //write to new txt file
    if (r.get("done")){tl="x ";if (r.get("donedate")!=null){tl+=new Date(r.get("donedate")).format("Y-m-d")+" ";}}
    if (r.get("priority")!=""){tl+="("+r.get("priority")+") ";}
    if (r.get("created")!=null){tl+=new Date(r.get("created")).format("Y-m-d")+" ";}
    tl+=r.get("task")+" ";
    tl+=r.get("tags")+" ";

    stream.WriteText(Trim(tl) ,1);

    //enter into Outlook if olid=""

}

//TODO: I think DUE: field from outlook items ahows wrong date
