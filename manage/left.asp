<!--#include file="conn.asp"-->
<!--#include file="denlu.asp"-->
<%
set rs=server.createobject("adodb.recordset")
rs.open "select quanxian from admin where zh='"&session("zh")&"'",conn,3
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">

<STYLE type=text/css>
.np
	{
	font-family: Webdings;
	font-size:18px;
	color:black;
	cursor:hand;
	}
A:link { COLOR: #20282F; FONT-SIZE: 12px; TEXT-DECORATION: none}
A:visited { COLOR: #20282F; FONT-SIZE: 12px; TEXT-DECORATION: none}
A:hover { COLOR: #FB8E01; FONT-SIZE: 12px; TEXT-DECORATION: none}
TD {	FONT-SIZE: 12px; line-height:150%}
</style>

<SCRIPT language=javascript>
function switchSysBar(){
	if (switchPoint.innerText==3){
		switchPoint.innerText=4
		document.all("mnuList").style.display="none"
		top.content.cols="16,*"
	}
	else{
		switchPoint.innerText=3
		document.all("mnuList").style.display=""				
		top.content.cols="162,*"		
	}
}	
</SCRIPT>

</HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" >
<SCRIPT>
function showitem(id,name)
{
	return ("<span><a target='main' href='"+id+"'>"+name+"</a></span><br>")
}

function switchoutlookBar(number)
{
	var i = outlookbar.opentitle;
	outlookbar.opentitle=number;
	var id1,id2,id1b,id2b
	if (number!=i && outlooksmoothstat==0){
	if (number!=-1)
		{
			if (i==-1)
				{
				id2="blankdiv";
				id2b="blankdiv";}
			else{
				id2="outlookdiv"+i;
				id2b="outlookdivin"+i;
				document.all("outlooktitle"+i).style.border="1px none navy";
				document.all("outlooktitle"+i).style.background="#006699";
				document.all("outlooktitle"+i).style.color="#ffffff";
				document.all("outlooktitle"+i).style.textalign="center";
				}
			id1="outlookdiv"+number
			id1b="outlookdivin"+number
			document.all("outlooktitle"+number).style.border="1px none white";
			document.all("outlooktitle"+number).style.background="#006699"; //title
			document.all("outlooktitle"+number).style.color="#ffffff";
			document.all("outlooktitle"+number).style.textalign="center";
			smoothout(id1,id2,id1b,id2b,0);
		}
	else
		{
			document.all("blankdiv").style.display="";
			document.all("blankdiv").sryle.height="100%";
			document.all("outlookdiv"+i).style.display="none";
			document.all("outlookdiv"+i).style.height="0%";
			document.all("outlooktitle"+i).style.border="1px none navy";
			document.all("outlooktitle"+i).style.background="#006699";
			document.all("outlooktitle"+i).style.color="#ffffff";
			document.all("outlooktitle"+i).style.textalign="center";
		}
	}
			
}

function smoothout(id1,id2,id1b,id2b,stat)
{
	if(stat==0){
		tempinnertext1=document.all(id1b).innerHTML;
		tempinnertext2=document.all(id2b).innerHTML;
		document.all(id1b).innerHTML="";
		document.all(id2b).innerHTML="";
		outlooksmoothstat=1;
		document.all(id1b).style.overflow="hidden";
		document.all(id2b).style.overflow="hidden";
		document.all(id1).style.height="0%";
		document.all(id1).style.display="";
		setTimeout("smoothout('"+id1+"','"+id2+"','"+id1b+"','"+id2b+"',"+outlookbar.inc+")",outlookbar.timedalay);
	}
	else
	{
		stat+=outlookbar.inc;
		if (stat>100)
			stat=100;
		document.all(id1).style.height=stat+"%";
		document.all(id2).style.height=(100-stat)+"%";
		if (stat<100) 
			setTimeout("smoothout('"+id1+"','"+id2+"','"+id1b+"','"+id2b+"',"+stat+")",outlookbar.timedalay);
		else
			{
			document.all(id1b).innerHTML=tempinnertext1;
			document.all(id2b).innerHTML=tempinnertext2;
			outlooksmoothstat=0;
			document.all(id1b).style.overflow="auto";
			document.all(id2).style.display="none";
			}
	}
}

function getOutLine()
{
	outline="<table "+outlookbar.otherclass+">";
	for (i=0;i<(outlookbar.titlelist.length);i++)
		{
			outline+="<tr><td name=outlooktitle"+i+" id=outlooktitle"+i+" ";		
			if (i!=outlookbar.opentitle) 
				outline+=" nowrap align=center style='cursor:hand;background-color:#006699;color:#ffffff;height:5;border:1 none navy' ";
			else
				outline+=" nowrap align=center style='cursor:hand;background-color:#006699;color:white;height:5;border:1 none white' ";
			outline+=outlookbar.titlelist[i].otherclass
			outline+=" onclick='switchoutlookBar("+i+")'><span class=smallFont>";
			outline+=outlookbar.titlelist[i].title+"</span></td></tr>";
			outline+="<tr><td name=outlookdiv"+i+" valign=top align=center id=outlookdiv"+i+" style='width:100%"
			if (i!=outlookbar.opentitle) 
				outline+=";display:none;height:0%;";
			else
				outline+=";display:;height:100%;";
			outline+="'><div name=outlookdivin"+i+" id=outlookdivin"+i+" style='overflow:auto;width:100%;height:100%'>";
			for (j=0;j<outlookbar.itemlist[i].length;j++)
				outline+=showitem(outlookbar.itemlist[i][j].key,outlookbar.itemlist[i][j].title);
			outline+="</div></td></tr>"
		}
	
	outline+="</table>"
	return outline

}

function show()
{
	var outline;
	outline="<div id=outLookBarDiv name=outLookBarDiv style='width=100%;height:100%'>"
	outline+=outlookbar.getOutLine();
	outline+="</div>"
	document.write(outline);
}

function theitem(intitle,instate,inkey)
{
	this.state=instate;
	this.otherclass=" nowrap ";
	this.key=inkey;
	this.title=intitle;
}

function addtitle(intitle)
{
	outlookbar.itemlist[outlookbar.titlelist.length]=new Array();
	outlookbar.titlelist[outlookbar.titlelist.length]=new theitem(intitle,1,0);
	return(outlookbar.titlelist.length-1);
}

function additem(intitle,parentid,inkey)
{
	if (parentid>=0 && parentid<=outlookbar.titlelist.length)
	{
		outlookbar.itemlist[parentid][outlookbar.itemlist[parentid].length]=new theitem(intitle,2,inkey);
		outlookbar.itemlist[parentid][outlookbar.itemlist[parentid].length-1].otherclass=" nowrap align=left style='height:5' ";
		return(outlookbar.itemlist[parentid].length-1);
	}
	else
		additem=-1;
}

function outlook()
{
	this.titlelist=new Array();
	this.itemlist=new Array();
	this.divstyle="style='height:100%;width:100%;overflow:auto' align=center";
	this.otherclass="border=0 cellspacing='0' cellpadding='0' style='height:100%;width:100%'valign=middle align=center ";
	this.addtitle=addtitle;
	this.additem=additem;
	this.starttitle=-1;
	this.show=show;
	this.getOutLine=getOutLine;
	this.opentitle=this.starttitle;
	this.reflesh=outreflesh;
	this.timedelay=50;
	this.inc=10;
	
}

function outreflesh()
{
	document.all("outLookBarDiv").innerHTML=outlookbar.getOutLine();
}

function locatefold(foldname)
{
	for (var i=0;i<outlookbar.titlelist.length;i++)
		if(foldname==outlookbar.titlelist[i].title)
			{
				 outlookbar.starttitle=i;
				 outlookbar.opentitle=i;
			}
	
}

var outlookbar=new outlook();
var tempinnertext1,tempinnertext2,outlooksmoothstat
outlooksmoothstat = 0;

</SCRIPT>

<SCRIPT>
var t;

<%
if instr(1,rs("quanxian"),",1,")>0 then
%>

t=outlookbar.addtitle('�����İ������')
outlookbar.additem('<br>�������',t,'Product_Cate_List.asp?ver=cn')
outlookbar.additem('<br>��Ʒ����',t,'Product_List.asp?ver=cn')
outlookbar.additem('<br>��Ϣ�������',t,'NewsCate_List.asp?ver=cn')
outlookbar.additem('<br>��Ϣ����',t,'News_list.asp?ver=cn')
outlookbar.additem('<br>��ҳ�õ�Ƭ',t,'banner_list.asp?ver=cn')
outlookbar.additem('<br>��ҵ��Ϣ',t,'company.asp?ver=cn')
outlookbar.additem('<br>��վ����',t,'site.asp?id=1')
outlookbar.additem('<br>����',t,'create_html.asp')
<%
end if
if instr(1,rs("quanxian"),",1,")>0 then
%>

t=outlookbar.addtitle('��Ӣ�İ������')
outlookbar.additem('<br>�������',t,'Product_Cate_List.asp?ver=en')
outlookbar.additem('<br>��Ʒ����',t,'Product_List.asp?ver=en')
outlookbar.additem('<br>��Ϣ�������',t,'NewsCate_List.asp?ver=en')
outlookbar.additem('<br>��Ϣ����',t,'News_list.asp?ver=en')
outlookbar.additem('<br>��ҳ�õ�Ƭ',t,'banner_list.asp?ver=en')
outlookbar.additem('<br>��ҵ��Ϣ',t,'company.asp?ver=en')
outlookbar.additem('<br>��վ����',t,'site.asp?id=2')
outlookbar.additem('<br>����',t,'create_en_html.asp')
<%
end if
if instr(1,rs("quanxian"),",3,")>0 then
%>
t=outlookbar.addtitle('��Ȩ�޹�����')
outlookbar.additem('<br>����Ա����',t,'quanxian.asp?pid=1')

<%
end if
%>
t=outlookbar.addtitle('��ϵͳ������')
outlookbar.additem('<br>�޸�����',t,'modipass.asp')
outlookbar.additem('<br>�˳�ϵͳ',t,'logout.asp')
</SCRIPT>
<table cellspacing=0 cellpadding=0 border=0 height="100%" width="162" align="center">
  <tr>
    <td valign="top">
		 <table border=0 align=center cellpadding=0 cellspacing=0 id=mnuList style="WIDTH: 140px; HEIGHT: 100%">
        <tr> 
          <td height=1></td>
        </tr>
        <tr> 
          <td height="100%" align=middle id=outLookBarShow name="outLookBarShow"> 
<script>
locatefold("��ϵͳ������")
outlookbar.show()	
</script> </td>
        </tr>
      </table>
	</td>
    <td width="16" bgcolor=#FFFFFF background="images/menu_bg_02.gif"> 
      <table cellspacing=0 cellpadding=0 border=0>
        <tr> 
          <td valign=center align=middle> <span class=np id=switchPoint onClick=switchSysBar() title="��/�رյ���">3</span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body></html>