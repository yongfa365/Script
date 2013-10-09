<!--
function ShowListPage(page,Pcount,TopicNum,maxperpage,strLink,ListName){
	var alertcolor = '#FF0000';
	maxperpage=Math.floor(maxperpage);
	TopicNum=Math.floor(TopicNum);
	page=Math.floor(page);
	var n,p;
	if ((page-1)%10==0) {
		p=(page-1) /10
	}else{
		p=(((page-1)-(page-1)%10)/10)
	}
	if(TopicNum%maxperpage==0) {
		n=TopicNum/maxperpage;
	}else{
		n=(TopicNum-TopicNum%maxperpage)/maxperpage+1;
	}
	document.write ('<table border="0" cellpadding="0" cellspacing="1" class="Tableborder5">');
	document.write ('<form method=post action="?pcount='+Pcount+strLink+'">');
	document.write ('<tr align="center">');
	document.write ('<td class="tabletitle1" title="'+ListName+'">&nbsp;'+ListName+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="总数">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="每页">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="页次">&nbsp;'+page+'/'+Pcount+'页&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page=1'+strLink+'" title="首页"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*10 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+p*10+strLink+'" title="上十页"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	if (page < 2){
		document.write ('<td class="tablebody1">&nbsp;首 页&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;上一页&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page=1'+strLink+'" title="首页">首 页</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+(page-1)+strLink+'" title="上一页">上一页</a>&nbsp;</td>');
	}
	if (Pcount-page < 1){
		document.write ('<td class="tablebody1">&nbsp;下一页&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;尾 页&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+(page+1)+strLink+'" title="下一页">下一页</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+Pcount+strLink+'" title="尾页">尾 页</a>&nbsp;</td>');
	}
	for (var i=p*10+1;i<p*10+11;i++){
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'" title="下十页"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+n+strLink+'" title="尾页"><font face=webdings>:</font></a>&nbsp;</td>');
	}
	document.write ('<td class="tablebody1"><input class="PageInput" type=text name="page" size=1 maxlength=10  value="'+page+'"></td>');
	document.write ('<td class="tablebody1"><input type=submit value=Go name=submit class="PageInput"></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
function ShowHtmlPage(page,Pcount,TopicNum,maxperpage,strLink,ExtName,ListName){
	var alertcolor = '#FF0000';
	maxperpage=Math.floor(maxperpage);
	TopicNum=Math.floor(TopicNum);
	page=Math.floor(page);
	var n,p;
	if ((page-1)%10==0) {
		p=(page-1) /10
	}else{
		p=(((page-1)-(page-1)%10)/10)
	}
	if(TopicNum%maxperpage==0) {
		n=TopicNum/maxperpage;
	}else{
		n=(TopicNum-TopicNum%maxperpage)/maxperpage+1;
	}
	document.write ('<table border="0" cellpadding="0" cellspacing="1" class="Tableborder5">');
	document.write ('<form method=post>');
	document.write ('<tr align="center">');
	document.write ('<td class="tabletitle1" title="'+ListName+'">&nbsp;'+ListName+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="总数">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="每页">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="页次">&nbsp;'+page+'/'+Pcount+'页&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="首页"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*10 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(p*10)+ExtName+'" title="上十页"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	if (page < 3){
		document.write ('<td class="tablebody1">&nbsp;首 页&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="上一页">上一页</a>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="首页">首 页</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(page-1)+ExtName+'" title="上一页">上一页</a>&nbsp;</td>');
	}
	if (Pcount-page < 1){
		document.write ('<td class="tablebody1">&nbsp;下一页&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;尾 页&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(page+1)+ExtName+'" title="下一页">下一页</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+Pcount+ExtName+'" title="尾页">尾 页</a>&nbsp;</td>');
	}
	for (var i=p*10+1;i<p*10+11;i++){
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+i+ExtName+'" title="下十页"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+n+ExtName+'" title="尾页"><font face=webdings>:</font></a>&nbsp;</td>');
	}
	//document.write ('<td class="tabletitle1" title="转到">&nbsp;GO&nbsp;</td>');
	document.write ('<td class="tablebody1"><select class="PageInput" name="page" size="1" onchange="javascript:window.location=this.options[this.selectedIndex].value;">');
	//document.write ('<option value="index'+ExtName+'">第1页</option>');
	for (var i=1;i<TopicNum;i++){
		if (i==page){
			document.write ('<option value="'+strLink+i+ExtName+'" selected>第'+i+'页</option>');
		}else{
			if (i==1){
				document.write ('<option value="index'+ExtName+'">第1页</option>');
			}else{
				document.write ('<option value="'+strLink+i+ExtName+'">第'+i+'页</option>');
			}
		}
		if (i==n) break;
	}
	document.write ('</select></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
//-->