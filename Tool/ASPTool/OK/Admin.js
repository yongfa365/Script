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
		if(confirm('ɾ������ѡ�У�'))
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

    //���¼�����onload���Ϊ�Ҳ�֪��JS���ֱ��д������ǲ��ǻ��ҳ��������ִ��
    //ʹ��<%=vd%>��ʽ���GridView��ID����ΪĳЩ����£���ʹ����MasterPage�������HTML��ID�ı仯
    //��ɫֵ�Ƽ�ʹ��Hex���� #f00 �� #ff0000
    //window.onload = function(){GridViewColor("trColor","#fff","#eee","#6df","#fd6");}
    
    //��������Ϊ�����������ָ��Ϊ��ֵ���򲻻ᷢ����Ӧ���¼�����
    //GridView ID, �����б���ɫ,�����б���ɫ,���ָ���б���ɫ,������󱳾�ɫ
    function GridViewColor(GridViewId, NormalColor, AlterColor, HoverColor, SelectColor){
        //��ȡ����Ҫ���Ƶ���
        if(!document.getElementById(GridViewId)){return;}
        var AllRows = document.getElementById(GridViewId).getElementsByTagName("tr");
        
        //����ÿһ�еı���ɫ���¼���ѭ����1��ʼ����0�����Աܿ���ͷ��һ��
        for(var i=1; i<AllRows.length; i++){
            //�趨����Ĭ�ϵı���ɫ
            AllRows[i].style.background = i%2==0?NormalColor:AlterColor;
            
            //���ָ�������ָ��ı���ɫ�������onmouseover/onmouseout�¼�
            //����ѡ��״̬���з����������¼�ʱ���ı���ɫ
            if(HoverColor != ""){
                AllRows[i].onmouseover = function(){if(!this.selected)this.style.background = HoverColor;}
                if(i%2 == 0){
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = NormalColor;}
                }
                else{
                    AllRows[i].onmouseout = function(){if(!this.selected)this.style.background = AlterColor;}
                }
            }

            //���ָ����������ı���ɫ�������onclick�¼�
            //���¼���Ӧ���޸ı�����е�ѡ��״̬
            if(SelectColor != ""){
                AllRows[i].onclick = function(){
                    this.style.background = this.style.background==SelectColor?HoverColor:SelectColor;
                    this.selected = !this.selected;
                }
            }
        }
    }