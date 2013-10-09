/*
 *ITgene前台JS表单脚本(输入&验证)库
 *版本:V1.1
 *作者:Mickey
 *时间:2006-3-9
 *版权:ITgene.cn
 *注:请使用本JS表单脚本(输入&验证)库的同时保留此版权信息，此是作者花了时间去收集整理以及编写的，谢谢!
 *	 此版本采用GB2312编码格式，请在使用前进行字符编码转换，以保证能够正常使用
 */

/*
 *以下是库函数目录及使用说明:
 *
 *常用
 *1、Trim=去除字符串前后空格									使用方法:String.trim()
 *2、ctrim=去除字符串中间空格									使用方法:String.ctrim()
 *3、onClickSelect=点中text框的时候，选中其中的文字				使用方法:在input位置加上 onClick/onFocus="onClickSelect();" 即可

 
 *
 *动态输入类													使用方法:在input位置加上 onkeypress="函数名" 即可
 *1、TextOnly=只允许输入字母、数字、下划线
 *2、TextNumOnly=只允许输入字母、数字
 *3、NumOnly=只允许输入数字
 *4、TelOnly=只能输入电话、"-"、"("、")"

 *
 *表单验证类
 *1、isAccount=是否帐号(由字母、数字、下划线组成){有两种选择,一种有长度限制}
 *2、isChinese=是否中文(由中文、数字、字母组成)
 *3、ismail=是否Email
 *4、isip=是否ip
 *5、PhoneCheck=电话号码检测(电话和手机)
 *6、isMobile=手机号码检测
 *7、isDate=是否短日期
 *8、isTime=是否时间
 *9、isDateTime=是否长日期

 *
 *其它函数
 *1、changeFrame=改变Frame大小
 *2、CheckAll=全选/全不选
 *3、onKeyDownDefault=回车->转->Tab
 *4、admin_Size=改变TextArea输入框高度

 *
 *其它验证正则表达式
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
 *Username : /^[a-z]\w{3,}$/i(用户名验证,带长度限制)
 */



//========================================================================常用函数
//--------------------------------除去前后空格
String.prototype.trim = function()
{
    //用正则表达式将前后空格用空字符串替代。
    return this.replace(/(^\s*)|(\s*$)/g, "");
}

//--------------------------------除去中间空格
String.prototype.ctrim = function()
{
    //用正则表达式将中间空格用空字符串替代。
	return this.replace(/\s/g,'');
}

//--------------------------------点中text框的时候，选中其中的文字
/**
*方法名：onClickSelect
*描述：点中text框的时候，选中其中的文字
*输入：空
*输出：空
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

//========================================================================动态输入类函数
//--------------------------------只允许输入字母、数字、下划线(动态判断)
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

//--------------------------------只允许输入字母、数字(动态判断)
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

//--------------------------------只允许输入数字(动态判断)
/**
*方法名：NumOnly()
*描述：只允许输入数字
*输入：空
*输出：空
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

//--------------------------------只能输入电话号码或者"－"或者"("或者")"
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
	//32=空格
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
//========================================================================动态输入函数(结束)


//========================================================================表单验证函数
//--------------------------------判断用户名（判断字符由字母和数字，下划线,点号组成.且开头的只能是下划线和字母）
function isAccount(str)
{
	if(/^[a-z]\w{3,}$/i.test(str))		 //用户名由字母和数字、下划线组成，且只能以字母开头，且长度最小为4位
	//if(/^([a-zA-z]{1})([\w]*)$/g.test(str))//用户名由字母和数字、下划线组成，且只能以字母开头
	{
		//alert(');
		return true;
	}
	else
		return false;
}

//--------------------------------判断只能输入中文、数字、字母
function isChinese(str)
{
	var pattern = /^[0-9a-zA-Z\u4e00-\u9fa5]+$/i;
	if (pattern.test(str))
	{
		return true;
	}
	else
	{
		//alert("只能包含中文、字母、数字");
		return false;
	}
}

//--------------------------------Email格式判断
function ismail(email)
{
	return(new RegExp(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/).test(email));
}

//--------------------------------IP格式判断
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

//--------------------------------判断电话号码/手机号码
function PhoneCheck(s) 
{
	var str=s;
	var reg=/(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/;
	//alert(reg.test(str));
	return reg.test(str);
}

//--------------------------------判断手机号码
function isMobile(str)
{      
	var reg=/^0{0,1}13[0-9]{9}$/;
	return reg.test(str);
}

//--------------------------------短日期(如2003-12-05)
function isDate(str)
{
	var r = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/); 
	if(r==null)
	{
		//alert('输入的信息不是日期格式（YYYY:MM:DD）'); 
		return false; 
	}
	if (r[1]<1 || r[3]<1 || r[3]-1>12 || r[4]<1 || r[4]>31)
	{
		//alert("日期格式（YYYY:MM:DD）不对");
		return false
	}
	var d= new Date(r[1], r[3]-1, r[4]); 
	return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]);
}

//--------------------------------短时间(如13:04:06)
function isTime(str)
{
	var a = str.match(/^(\d{1,2})(:)?(\d{1,2})\2(\d{1,2})$/);
	if (a == null) 
	{
		//alert('输入的信息不是时间格式（HH:MM:SS）'); 
		return false;
	}
	if (a[1]>23 || a[1]<0 || a[3]>60 || a[3]<0 || a[4]>60 || a[4]<0)
	{
		//alert("时间格式（0<=HH<23:0<=MM<60:0<=SS<60）不对");
		return false
	}
	return true;
}

//--------------------------------长时间(如2003-12-05 13:04:06)
function isDateTime(str)
{
	var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/;
	var r = str.match(reg); 
	if(r==null)
	{
		//alert('输入的信息不是时间格式（YYYY-MM-DD HH:MM:SS）'); 
		return false;
	}
	var d= new Date(r[1], r[3]-1,r[4],r[5],r[6],r[7]); 
	return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]&&d.getHours()==r[5]&&d.getMinutes()==r[6]&&d.getSeconds()==r[7]);
}

//========================================================================表单验证函数(结束)

//========================================================================其它函数
//用于改变当前后台框架的尺寸(参数:对象,关闭时显示图片,打开时显示图片,关闭显示尺寸,打开显示尺寸)
//							例:this,'../images/openout.gif','../images/setby.gif','top.frame.cols=\'0,*\'','top.frame.cols=\'150,*\''
function changeFrame(ob,s1,s2,evalString1,evalString2)
{
	if (ob.alt == "收起")
	{
		ob.alt = "展开";
		ob.src = s1;
		eval(evalString1);
	}
	else
	{
		ob.alt = "收起";
		ob.src = s2;
		eval(evalString2);
	}
}

//用于全选/全不选(参数form=表单)
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

//--------------------------------用户在button以外的页面元素中按回车-->转换-->按tab键
/**
*方法名：onKeyDownDefault
*描述：用户在button以外的页面元素中按回车，转换为按tab键
*输入：空
*输出：空
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

// 修改编辑栏高度
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

//取得字符串中中文/汉字的个数
function getChineseNum(obstring)
{
	var pattern = /^[\u4e00-\u9fa5]+$/i;
	var maxL,minL;
	maxL = obstring.length;				//原始长度
	obstring = obstring.replace(pattern,"");
	minL = obstring.length;				//处理后的长度
	return (maxL - minL)
}
//========================================================================其它函数(结束)