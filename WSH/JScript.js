/*
 *ITgeneǰ̨JS���ű�(����&��֤)��
 *�汾:V1.1
 *����:Mickey
 *ʱ��:2006-3-9
 *��Ȩ:ITgene.cn
 *ע:��ʹ�ñ�JS���ű�(����&��֤)���ͬʱ�����˰�Ȩ��Ϣ���������߻���ʱ��ȥ�ռ������Լ���д�ģ�лл!
 *	 �˰汾����GB2312�����ʽ������ʹ��ǰ�����ַ�����ת�����Ա�֤�ܹ�����ʹ��
 */

/*
 *�����ǿ⺯��Ŀ¼��ʹ��˵��:
 *
 *����
 *1��Trim=ȥ���ַ���ǰ��ո�									ʹ�÷���:String.trim()
 *2��ctrim=ȥ���ַ����м�ո�									ʹ�÷���:String.ctrim()
 *3��onClickSelect=����text���ʱ��ѡ�����е�����				ʹ�÷���:��inputλ�ü��� onClick/onFocus="onClickSelect();" ����

 
 *
 *��̬������													ʹ�÷���:��inputλ�ü��� onkeypress="������" ����
 *1��TextOnly=ֻ����������ĸ�����֡��»���
 *2��TextNumOnly=ֻ����������ĸ������
 *3��NumOnly=ֻ������������
 *4��TelOnly=ֻ������绰��"-"��"("��")"

 *
 *����֤��
 *1��isAccount=�Ƿ��ʺ�(����ĸ�����֡��»������){������ѡ��,һ���г�������}
 *2��isChinese=�Ƿ�����(�����ġ����֡���ĸ���)
 *3��ismail=�Ƿ�Email
 *4��isip=�Ƿ�ip
 *5��PhoneCheck=�绰������(�绰���ֻ�)
 *6��isMobile=�ֻ�������
 *7��isDate=�Ƿ������
 *8��isTime=�Ƿ�ʱ��
 *9��isDateTime=�Ƿ�����

 *
 *��������
 *1��changeFrame=�ı�Frame��С
 *2��CheckAll=ȫѡ/ȫ��ѡ
 *3��onKeyDownDefault=�س�->ת->Tab
 *4��admin_Size=�ı�TextArea�����߶�

 *
 *������֤������ʽ
 *Email : /^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/
 *Phone : /^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$/
 *Url : /^http:\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\"\"])*$/
 *Currency : /^\d+(\.\d+)?$/
 *Number : /^\d+$/
 *Zip : /^[1-9]\d{5}$/
 *QQ : /^[1-9]\d{4,8}$/
 *Integer : /^[-\+]?\d+$/
 *Double : /^[-\+]?\d+(\.\d+)?$/
 *English : /^[A-Za-z]+$/
 *Chinese :  /^[\u0391-\uFFE5]+$/
 *UnSafe : /^(([A-Z]*|[a-z]*|\d*|[-_\~!@#\$%\^&\*\.\(\)\[\]\{\}<>\?\\\/\'\"]*)|.{0,5})$|\s/
 *Username : /^[a-z]\w{3,}$/i(�û�����֤,����������)
 */



//========================================================================���ú���
//--------------------------------��ȥǰ��ո�
String.prototype.trim = function()
{
    //��������ʽ��ǰ��ո��ÿ��ַ��������
    return this.replace(/(^\s*)|(\s*$)/g, "");
}

//--------------------------------��ȥ�м�ո�
String.prototype.ctrim = function()
{
    //��������ʽ���м�ո��ÿ��ַ��������
	return this.replace(/\s/g,'');
}

//--------------------------------����text���ʱ��ѡ�����е�����
/**
*��������onClickSelect
*����������text���ʱ��ѡ�����е�����
*���룺��
*�������
**/
function onClickSelect()
{
    var obj = document.activeElement;
    if(obj.tagName == "TEXTAREA")
	{
        obj.select();
	}
    if(obj.tagName == "INPUT" )
	{
        if(obj.type == "text")
			obj.select();
	}
}

//========================================================================��̬�����ຯ��
//--------------------------------ֻ����������ĸ�����֡��»���(��̬�ж�)
function TextOnly(){
    var i= window.event.keyCode
	//8=backspace
	//9=tab
	//37=left arrow
	//39=right arrow
	//46=delete
	//48~57=0~9
	//97~122=a~z
	//65~90=A~Z
	//95=_
    if (!((i<=57 && i>=48)||(i>=97 && i<=122)||(i>=65 && i<=90)||(i==95)||(i==8)||(i==9)||(i==37)||(i==39)||(i==46)))
	{
        //window.event.keyCode=27;
		event.returnValue=false
		return false;
    }
	else
	{
        //window.event.keyCode=keycode;
		return true;
    }
}

//--------------------------------ֻ����������ĸ������(��̬�ж�)
function TextNumOnly(){
    var i= window.event.keyCode
	//8=backspace
	//9=tab
	//37=left arrow
	//39=right arrow
	//46=delete
	//48~57=0~9
	//97~122=a~z
	//65~90=A~Z
	//95=_
    if (!((i<=57 && i>=48)||(i>=97 && i<=122)||(i>=65 && i<=90)||(i==8)||(i==9)||(i==37)||(i==39)||(i==46)))
	{
        //window.event.keyCode=27;
		event.returnValue=false
		return false;
    }
	else
	{
        //window.event.keyCode=keycode;
		return true;
    }
}

//--------------------------------ֻ������������(��̬�ж�)
/**
*��������NumOnly()
*������ֻ������������
*���룺��
*�������
**/
function NumOnly(){
    var i= window.event.keyCode
	//8=backspace
	//9=tab
	//37=left arrow
	//39=right arrow
	//46=delete
	//48~57=0~9
    if ((i>57 || i<48) && (i!=8) && (i!=9) && (i!= 37) && (i!=39) && (i!=46))
	{
        window.event.keyCode=27;
		return false;
    }
	else
	{
        //window.event.keyCode=keycode;
		return true;
    }
}

//--------------------------------ֻ������绰�������"��"����"("����")"
function TelOnly(){
    var i= window.event.keyCode
	//8=backspace
	//9=tab
	//37=left arrow
	//39=right arrow
	//46=delete
	//48~57=0~9
	//40=(
	//41=)
	//45=-
	//32=�ո�
    if ((i>57 || i<48) && (i!=8) && (i!=9) && (i!= 37) && (i!=39) && (i!=46) && (i!=40) && (i!=41) && (i!=45) && (i!=32))
	{
        window.event.keyCode=27;
		return false;
    }
	else
	{
        //window.event.keyCode=keycode;
		return true;
    }
}
//========================================================================��̬���뺯��(����)


//========================================================================����֤����
//--------------------------------�ж��û������ж��ַ�����ĸ�����֣��»���,������.�ҿ�ͷ��ֻ�����»��ߺ���ĸ��
function isAccount(str)
{
	if(/^[a-z]\w{3,}$/i.test(str))		 //�û�������ĸ�����֡��»�����ɣ���ֻ������ĸ��ͷ���ҳ�����СΪ4λ
	//if(/^([a-zA-z]{1})([\w]*)$/g.test(str))//�û�������ĸ�����֡��»�����ɣ���ֻ������ĸ��ͷ
	{
		//alert(');
		return true;
	}
	else
		return false;
}

//--------------------------------�ж�ֻ���������ġ����֡���ĸ
function isChinese(str)
{
	var pattern = /^[0-9a-zA-Z\u4e00-\u9fa5]+$/i;
	if (pattern.test(str))
	{
		return true;
	}
	else
	{
		//alert("ֻ�ܰ������ġ���ĸ������");
		return false;
	}
}

//--------------------------------Email��ʽ�ж�
function ismail(email)
{
	return(new RegExp(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/).test(email));
}

//--------------------------------IP��ʽ�ж�
function isip(s)
{
	var check=function(v)
	{
		try
		{
			return (v<=255 && v>=0)
		}
		catch(x)
		{
			return false
		}
	};
	var re=s.split(".")
	return (re.length==4)?(check(re[0]) && check(re[1]) && check(re[2]) && check(re[3])):false;
}

//--------------------------------�жϵ绰����/�ֻ�����
function PhoneCheck(s) 
{
	var str=s;
	var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/;
	//alert(reg.test(str));
	return reg.test(str);
}

//--------------------------------�ж��ֻ�����
function isMobile(str)
{      
	var reg=/^0{0,1}13[0-9]{9}$/;
	return reg.test(str);
}

//--------------------------------������(��2003-12-05)
function isDate(str)
{
	var r = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/); 
	if(r==null)
	{
		//alert('�������Ϣ�������ڸ�ʽ��YYYY:MM:DD��'); 
		return false; 
	}
	if (r[1]<1 || r[3]<1 || r[3]-1>12 || r[4]<1 || r[4]>31)
	{
		//alert("���ڸ�ʽ��YYYY:MM:DD������");
		return false
	}
	var d= new Date(r[1], r[3]-1, r[4]); 
	return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]);
}

//--------------------------------��ʱ��(��13:04:06)
function isTime(str)
{
	var a = str.match(/^(\d{1,2})(:)?(\d{1,2})\2(\d{1,2})$/);
	if (a == null) 
	{
		//alert('�������Ϣ����ʱ���ʽ��HH:MM:SS��'); 
		return false;
	}
	if (a[1]>23 || a[1]<0 || a[3]>60 || a[3]<0 || a[4]>60 || a[4]<0)
	{
		//alert("ʱ���ʽ��0<=HH<23:0<=MM<60:0<=SS<60������");
		return false
	}
	return true;
}

//--------------------------------��ʱ��(��2003-12-05 13:04:06)
function isDateTime(str)
{
	var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/;
	var r = str.match(reg); 
	if(r==null)
	{
		//alert('�������Ϣ����ʱ���ʽ��YYYY-MM-DD HH:MM:SS��'); 
		return false;
	}
	var d= new Date(r[1], r[3]-1,r[4],r[5],r[6],r[7]); 
	return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]&&d.getHours()==r[5]&&d.getMinutes()==r[6]&&d.getSeconds()==r[7]);
}

//========================================================================����֤����(����)

//========================================================================��������
//���ڸı䵱ǰ��̨��ܵĳߴ�(����:����,�ر�ʱ��ʾͼƬ,��ʱ��ʾͼƬ,�ر���ʾ�ߴ�,����ʾ�ߴ�)
//							��:this,'../images/openout.gif','../images/setby.gif','top.frame.cols=\'0,*\'','top.frame.cols=\'150,*\''
function changeFrame(ob,s1,s2,evalString1,evalString2)
{
	if (ob.alt == "����")
	{
		ob.alt = "չ��";
		ob.src = s1;
		eval(evalString1);
	}
	else
	{
		ob.alt = "����";
		ob.src = s2;
		eval(evalString2);
	}
}

//����ȫѡ/ȫ��ѡ(����form=��)
function CheckAll(form)     
{
  for (var i=0;i<form.length;i++)     
    {     
		var e = form[i];
		if (e.type=="checkbox")
		{
			if (e.name!="AllChoice")
			{
				if (form.AllChoice.checked==true)
				{
					if (e.disabled==false) e.checked=true;

				}
				else
				{
					if (e.disabled==false) e.checked=false;
				}
			}
		}
    }     
}

//--------------------------------�û���button�����ҳ��Ԫ���а��س�-->ת��-->��tab��
/**
*��������onKeyDownDefault
*�������û���button�����ҳ��Ԫ���а��س���ת��Ϊ��tab��
*���룺��
*�������
**/
function onKeyDownDefault()
{
    if (window.event.keyCode == 13 && window.event.ctrlKey == false && window.event.altKey == false){
        if (window.event.srcElement.type != "button")
            window.event.keyCode = 9;
    }
    else{
        return true;
    }
}

// �޸ı༭���߶�
function admin_Size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0)
	{
		obj.width="90%";
	}
}

//ȡ���ַ���������/���ֵĸ���
function getChineseNum(obstring)
{
	var pattern = /^[\u4e00-\u9fa5]+$/i;
	var maxL,minL;
	maxL = obstring.length;				//ԭʼ����
	obstring = obstring.replace(pattern,"");
	minL = obstring.length;				//�����ĳ���
	return (maxL - minL)
}
//========================================================================��������(����)