function DelChkID(delurl,id)
{
	var obj=document.getElementById(id).getElementsByTagName("input");
	var ids=""
	for(i=1;i<obj.length;i++)
	{
		if(obj[i].type=="checkbox" && obj[i].checked)
		{
			ids += "," + obj[i].value;
		}
	}
	if(ids.length>0)
	{
		if(confirm('删除所有选中？'))
		{
			window.location=delurl + ids.substring(1) + "&ModiURL=" + escape(window.location);
		}
	}
}

function CheckBox(id)
{
	var obj=document.getElementById(id).getElementsByTagName("input");
	for(i=1;i<obj.length;i++)
	{
		if(obj[i].type=="checkbox")
		{
			obj[i].checked=obj[i].checked==false?true:false;
		}
	}
}

    //把事件放在onload里，因为我不知道JS如果直接写到这儿是不是会等页面加载完才执行
    //使用<%=vd%>方式输出GridView的ID是因为某些情况下（如使用了MasterPage）会造成HTML中ID的变化
    //颜色值推荐使用Hex，如 #f00 或 #ff0000
    //window.onload = function(){GridViewColor("trColor","#fff","#eee","#6df","#fd6");}
    
    //参数依次为（后两个如果指定为空值，则不会发生相应的事件）：
    //GridView ID, 正常行背景色,交替行背景色,鼠标指向行背景色,鼠标点击后背景色
    function GridViewColor(GridViewId, NormalColor, AlterColor, HoverColor, SelectColor){
        //获取所有要控制的行
        if(!document.getElementById(GridViewId)){return;}
        var AllRows = document.getElementById(GridViewId).getElementsByTagName("tr");
        
        //设置每一行的背景色和事件，循环从1开始而非0，可以避开表头那一行
        for(var i=1; i<AllRows.length; i++){
            //设定本行默认的背景色
            AllRows[i].style.background = i%2==0?NormalColor:AlterColor;
            
            //如果指定了鼠标指向的背景色，则添加onmouseover/onmouseout事件
            //处于选中状态的行发生这两个事件时不改变颜色
            if(HoverColor != ""){
                AllRows[i].onmouseover = function(){if(!this.selected)this.style.background = HoverColor;}
                if(i%2 == 0){
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = NormalColor;}
                }
                else{
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = AlterColor;}
                }
            }

            //如果指定了鼠标点击的背景色，则添加onclick事件
            //在事件响应中修改被点击行的选中状态
            if(SelectColor != ""){
                AllRows[i].onclick = function(){
                    this.style.background = this.style.background==SelectColor?HoverColor:SelectColor;
                    this.selected = !this.selected;
                }
            }
        }
    }