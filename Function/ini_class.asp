包括写入ini   
  呵呵   
  <%   
  'Power   By   Tim   
  '文件摘要：INI类   
  '文件版本：3.0     
  '文本创建日期：2:17   2004-12-14   
  '=================   属性说明   ================   
  'INI.OpenFile   =   文件路径（使用虚拟路径需在外部定义）   
  'INI.CodeSet     =   编码设置，默认为   GB2312   
  'INI.IsTrue       =   检测文件是否正常（存在）   
  '================   方法说明   =================   
  'IsGroup(组名)                         检测组是否存在   
  'IsNode(组名,节点名)                         检测节点是否存在   
  'GetGroup(组名)                         读取组信息   
  'CountGroup()                         统计组数量   
  'ReadNode(组名,节点名)                         读取节点数据   
  'WriteGroup(组名)                         创建组   
  'WriteNode(组,节点,节点数据)             插入/更新节点数据   
  'DeleteGroup(组名)                         删除组   
  'DeleteNode(组名,节点名)             删除节点   
  'Save()                                     保存文件   
  'Close()                                     清除内部数据（释放）   
  '===============================================   
    
    
  Class   INI_Class   
  '===============================================   
            Private   Stream                         '//   Stream   对象   
            Private   FilePath             '//   文件路径   
            Public   Content                         '//   文件数据   
            Public   IsTrue                         '//   文件是否存在   
            Public   IsAnsi                         '//   记录是否二进制   
            Public   CodeSet                         '//   数据编码   
  '================================================   
              
            '//   初始化   
            Private   Sub   Class_Initialize()   
                        Set   Stream             =   Server.CreateObject("ADODB.Stream")   
                        Stream.Mode             =   3   
                        Stream.Type             =   2   
                        CodeSet                         =   "gb2312"   
                        IsAnsi                         =   True   
                        IsTrue                         =   True   
            End   Sub   
              
              
            '//   二进制流转换为字符串   
            Private   Function   Bytes2bStr(bStr)   
                        if   Lenb(bStr)=0   Then   
                                    Bytes2bStr   =   ""   
                                    Exit   Function   
                        End   if   
                          
                        Dim   BytesStream,StringReturn   
                        Set   BytesStream   =   Server.CreateObject("ADODB.Stream")   
                        With   BytesStream   
                                    .Type                 =   2   
                                    .Open   
                                    .WriteText       bStr   
                                    .Position         =   0   
                                    .Charset           =   CodeSet   
                                    .Position         =   2   
                                    StringReturn   =   .ReadText   
                                    .Close   
                        End   With   
                        Bytes2bStr               =   StringReturn   
                        Set   BytesStream               =   Nothing   
                        Set   StringReturn   =   Nothing   
            End   Function   
              
              
            '//   设置文件路径   
            Property   Let   OpenFile(INIFilePath)   
                        FilePath   =   INIFilePath   
                        Stream.Open   
                        On   Error   Resume   Next   
                        Stream.LoadFromFile(FilePath)   
                        '//   文件不存在时返回给   IsTrue   
                        if   Err.Number<>0   Then   
                                    IsTrue   =   False   
                                    Err.Clear   
                        End   if   
                        Content   =   Stream.ReadText(Stream.Size)   
                        if   Not   IsAnsi   Then   Content=Bytes2bStr(Content)   
            End   Property   
              
              
            '//   检测组是否存在[参数:组名]   
            Public   Function   IsGroup(GroupName)   
                        if   Instr(Content,"["&GroupName&"]")>0   Then   
                                    IsGroup   =   True   
                        Else   
                                    IsGroup   =   False   
                        End   if   
            End   Function   
              
              
            '//   读取组信息[参数:组名]   
            Public   Function   GetGroup(GroupName)   
                        Dim   TempGroup   
                        if   Not   IsGroup(GroupName)   Then   Exit   Function   
                        '//   开始寻找头部截取   
                        TempGroup   =   Mid(Content,Instr(Content,"["&GroupName&"]"),Len(Content))   
                        '//   剔除尾部   
                        if   Instr(TempGroup,VbCrlf&"[")>0   Then   TempGroup=Left(TempGroup,Instr(TempGroup,VbCrlf&"[")-1)   
                        if   Right(TempGroup,1)<>Chr(10)   Then   TempGroup=TempGroup&VbCrlf   
                        GetGroup   =   TempGroup   
            End   Function   
              
              
            '//   检测节点是否存在[参数:组名,节点名]   
            Public   Function   IsNode(GroupName,NodeName)   
                        if   Instr(GetGroup(GroupName),NodeName&"=")   Then   
                                    IsNode   =   True   
                        Else   
                                    IsNode   =   False   
                        End   if   
            End   Function   
              
              
            '//   创建组[参数:组名]   
            Public   Sub   WriteGroup(GroupName)   
                        if   Not   IsGroup(GroupName)   And   GroupName<>""   Then   
                                    Content   =   Content   &   "["   &   GroupName   &   "]"   &   VbCrlf   
                        End   if   
            End   Sub   
              
              
            '//   读取节点数据[参数:组名,节点名]   
            Public   Function   ReadNode(GroupName,NodeName)   
                        if   Not   IsNode(GroupName,NodeName)   Then   Exit   Function   
                        Dim   TempContent   
                        '//   取组信息   
                        TempContent   =   GetGroup(GroupName)   
                        '//   取当前节点数据   
                        TempContent   =   Right(TempContent,Len(TempContent)-Instr(TempContent,NodeName&"=")+1)   
                        TempContent   =   Replace(Left(TempContent,Instr(TempContent,VbCrlf)-1),NodeName&"=","")   
                        ReadNode   =   ReplaceData(TempContent,0)   
            End   Function   
              
              
            '//   写入节点数据[参数:组名,节点名,节点数据]   
            Public   Sub   WriteNode(GroupName,NodeName,NodeData)   
                        '//   组不存在时写入组   
                        if   Not   IsGroup(GroupName)   Then   WriteGroup(GroupName)   
                          
                        '//   寻找位置插入数据   
                        '///   获取组   
                        Dim   TempGroup   :   TempGroup   =   GetGroup(GroupName)   
                          
                        '///   在组尾部追加   
                        Dim   NewGroup   
                        if   IsNode(GroupName,NodeName)   Then   
                                    NewGroup   =   Replace(TempGroup,NodeName&"="&ReplaceData(ReadNode(GroupName,NodeName),1),NodeName&"="&ReplaceData(NodeData,1))   
                        Else   
                                    NewGroup   =   TempGroup   &   NodeName   &   "="   &   ReplaceData(NodeData,1)   &   VbCrlf   
                        End   if   
                          
                        Content   =   Replace(Content,TempGroup,NewGroup)   
            End   Sub   
              
              
            '//   删除组[参数:组名]   
            Public   Sub   DeleteGroup(GroupName)   
                        Content   =   Replace(Content,GetGroup(GroupName),"")   
            End   Sub   
              
              
            '//   删除节点[参数:组名,节点名]   
            Public   Sub   DeleteNode(GroupName,NodeName)   
                        Dim   TempGroup   
                        Dim   NewGroup   
                        TempGroup   =   GetGroup(GroupName)   
                        NewGroup   =   Replace(TempGroup,NodeName&"="&ReadNode(GroupName,NodeName)&VbCrlf,"")   
                        if   Right(NewGroup,1)<>Chr(10)   Then   NewGroup   =   NewGroup&VbCrlf   
                        Content   =   Replace(Content,TempGroup,NewGroup)   
            End   Sub   
              
              
            '//   替换字符[实参:替换目标,数据流向方向]   
            '             字符转换[防止关键符号出错]   
            '             [                         --->             {(@)}   
            '             ]                         --->             {(#)}   
            '             =                         --->             {($)}   
            '             回车             --->             {(1310)}   
            Public   Function   ReplaceData(Data_Str,IsIn)   
                        if   IsIn   Then   
                                    ReplaceData   =   Replace(Replace(Replace(Data_Str,"[","{(@)}"),"]","{(#)}"),"=","{($)}")   
                                    ReplaceData   =   Replace(ReplaceData,Chr(13)&Chr(10),"{(1310)}")   
                        Else   
                                    ReplaceData   =   Replace(Replace(Replace(Data_Str,"{(@)}","["),"{(#)}","]"),"{($)}","=")   
                                    ReplaceData   =   Replace(ReplaceData,"{(1310)}",Chr(13)&Chr(10))   
                        End   if   
            End   Function   
              
              
            '//   保存文件数据   
            Public   Sub   Save()   
                        With   Stream   
                                    .Close   
                                    .Open   
                                    .WriteText   Content   
                                    .SaveToFile   FilePath,2   
                        End   With   
            End   Sub   
              
              
            '//   关闭、释放   
            Public   Sub   Close()   
                        Set   Stream   =   Nothing   
                        Set   Content   =   Nothing   
            End   Sub   
              
  End   Class   
    
    
  Dim   INI   
  Set   INI   =   New   INI_Class   
  INI.OpenFile   =   Server.MapPath("Config.ini")   
  '==========   这是写入ini数据   ==========   
  INI.WriteNode("SiteConfig","SiteName","Leadbbs极速论坛")   
  INI.WriteNode("SiteConfig","Mail","leadbbs@leadbbs.com")   
  INI.Save()   
  '==========   这是读取ini数据   ==========   
  Response.Write("站点名称："INI.ReadNode("SiteConfig","SiteName"))   
  %> 