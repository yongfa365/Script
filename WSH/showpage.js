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
	document.write ('<td class="tabletitle1" title="����">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="ÿҳ">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="ҳ��">&nbsp;'+page+'/'+Pcount+'ҳ&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page=1'+strLink+'" title="��ҳ"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*10 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+p*10+strLink+'" title="��ʮҳ"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	if (page < 2){
		document.write ('<td class="tablebody1">&nbsp;�� ҳ&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;��һҳ&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page=1'+strLink+'" title="��ҳ">�� ҳ</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+(page-1)+strLink+'" title="��һҳ">��һҳ</a>&nbsp;</td>');
	}
	if (Pcount-page < 1){
		document.write ('<td class="tablebody1">&nbsp;��һҳ&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;β ҳ&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+(page+1)+strLink+'" title="��һҳ">��һҳ</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+Pcount+strLink+'" title="βҳ">β ҳ</a>&nbsp;</td>');
	}
	for (var i=p*10+1;i<p*10+11;i++){
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'" title="��ʮҳ"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+n+strLink+'" title="βҳ"><font face=webdings>:</font></a>&nbsp;</td>');
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
	document.write ('<td class="tabletitle1" title="����">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="ÿҳ">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="ҳ��">&nbsp;'+page+'/'+Pcount+'ҳ&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="��ҳ"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*10 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(p*10)+ExtName+'" title="��ʮҳ"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	if (page < 3){
		document.write ('<td class="tablebody1">&nbsp;�� ҳ&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="��һҳ">��һҳ</a>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="index'+ExtName+'" title="��ҳ">�� ҳ</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(page-1)+ExtName+'" title="��һҳ">��һҳ</a>&nbsp;</td>');
	}
	if (Pcount-page < 1){
		document.write ('<td class="tablebody1">&nbsp;��һҳ&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;β ҳ&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+(page+1)+ExtName+'" title="��һҳ">��һҳ</a>&nbsp;</td>');
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+Pcount+ExtName+'" title="βҳ">β ҳ</a>&nbsp;</td>');
	}
	for (var i=p*10+1;i<p*10+11;i++){
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+i+ExtName+'" title="��ʮҳ"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+n+ExtName+'" title="βҳ"><font face=webdings>:</font></a>&nbsp;</td>');
	}
	//document.write ('<td class="tabletitle1" title="ת��">&nbsp;GO&nbsp;</td>');
	document.write ('<td class="tablebody1"><select class="PageInput" name="page" size="1" onchange="javascript:window.location=this.options[this.selectedIndex].value;">');
	//document.write ('<option value="index'+ExtName+'">��1ҳ</option>');
	for (var i=1;i<TopicNum;i++){
		if (i==page){
			document.write ('<option value="'+strLink+i+ExtName+'" selected>��'+i+'ҳ</option>');
		}else{
			if (i==1){
				document.write ('<option value="index'+ExtName+'">��1ҳ</option>');
			}else{
				document.write ('<option value="'+strLink+i+ExtName+'">��'+i+'ҳ</option>');
			}
		}
		if (i==n) break;
	}
	document.write ('</select></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
//-->