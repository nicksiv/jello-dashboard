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
           {name: 'olid'},
]);


function todoTxtSync()
{
//Todotxt sync main function


    var mp=Ext.getCmp("mainpanel2");
    mp.getEl().update("");

  
var simple = new Ext.FormPanel({
        frame:true,
    	title:'Todo.txt file sync',
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

            {
                fieldLabel: 'Skip tags (comma separated)',
                id: 'twskip',
                visible:true,
                value:'',
                tabIndex:3,
		width:420,
                allowBlank:true
            },
        new Ext.form.Label({
				html:'',
				width:650,
	    			id:'reslabel2',
				cls:'formfilelist',
            			height:50
				})
	    ,new Ext.Button({text:'<b>Sync Todo.txt File</b>',id:'twgo',fieldLabel:' ',autoWidth:false,width:180,height:30,labelSeparator:'',listeners:{click:function(p){todoSync();}}})

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
    var txtfile=fso.GetFile(fname);}catch(e){if (!confirm("File "+fname+" does not exist \n Do you want to create it?")){return;}else
					  {  
					     
					      try{fso.CreateFolder(fso.GetFileFolder(fname));}catch(g){}
					     try{
					      fso.CreateTextFile(fname);
					      var txtfile=fso.GetFile(fname); }catch(f){alert("Something went wrong. File could not be created!\n"+f.description);return;}
					  }}

var txtLines=new Array();

//open the todo.txt file and get lines
var ts = txtfile.OpenAsTextStream(1,-2);
var s="";
    do
	{
	    s=ts.ReadLine();
	    txtLines.push(s);
	} while (!ts.AtEndOfStream)	

ts.close();


    //01. read each line, handle and add to a table of values
    for (var x=0;x<txtLines.length;x++)
    {handleTxtLine(txtLines[x],x);}
    
    //02. read all outlook jello items, compare and add to the table
    var iF=NSpace.GetFolderFromID(jello.actionFolder).Items;
    var dasl="http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003 <> 2"
    var its=iF.Restrict("@SQL="+dasl);
    var counter=its.Count;
    	for (var y=1;y<=counter;y++)
	    {
		handleOLItemToTxt(its(y));
	    }
alert(txtStore.getCount());

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


function handleTxtLine(tLine,num)
{
//add fields of txt line to table

//split the todotxt line in spaces
var items=tLine.split(" ");

var fpri="";
var ftask="";
var fdone=false;
var fdonedate="";
var fcreated="";
var ftags="";
var ftask="";

    //priority item 
    if (items[0]=="("){fpri=items[0].substr(1,1);items[0]="";}
    	
    //done item
    if (items[0]=="x")
    {fdone=true;fdonedate=items[1];
     if (isDate(items[2])){fcreated=items[2];}
     items[0]="";items[1]="";items[2]="";}

    //simple task with creation date
    if (isDate(items[0])){fcreated=items[0];items[0]="";}
    
    //tags
    for (var x=0;x<items.length;x++)
	{
	    if (items[x].substr(0,1)=="@" || items[x].substr(0,1)=="+")
		{ftags+=items[x]+" ";items[x]="";}
	}

    ftask=items.join(" ");

var tr=new txtRecord({
priority:fpri,
done:fdone,
donedate:Date.parseDate(fdonedate,"Y-m-d"),
created:Date.parseDate(fcreated,"Y-m-d"),
tags:ftags,
task:Trim(ftask),
lineid:num
});
txtStore.add(tr);


}


function JelloTagsToTxt(cats)
{
var tTags="";
var nextFlag=false;
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

    var ix=txtStore.find("task",new RegExp("^"+olitem.Subject+"$"));
    var r=txtStore.getAt(ix);
    	if (r==null)
	 {
	     //outlook item not found in txt. Add it
	     fpri="";fcreated=olitem.CreationTime;
	     ftags=JelloTagsToTxt(olitem.itemProperties.item(catProperty).Value);
	     ftask=Trim(olitem.Subject);
	     var opri=getJPriorityProperty(olitem.EntryID);
	     if (opri=="1"){fpri="B";}
	     if (opri=="2"){fpri="C";}
	     if (opri=="3"){fpri="D";}
	     if (ftags.substr(0,7)=="{{{*}}}"){fpri="A";ftags=ftags.substr(7,ftags.length);}
	     var tr=new txtRecord({
		 priority:fpri,
		 //done:fdone,
		 //donedate:Date.parseDate(fdonedate,"Y-m-d"),
		 created:fcreated,
		 tags:ftags,
		 task:Trim(ftask),
		 //lineid:num
	     	});
	     txtStore.add(tr);

	 }
    	else
	 {
	     //outlook item found in txt. Need to spot differences?
	 }
}
