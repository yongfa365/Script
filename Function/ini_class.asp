����д��ini   
  �Ǻ�   
  <%   
  'Power   By   Tim   
  '�ļ�ժҪ��INI��   
  '�ļ��汾��3.0     
  '�ı��������ڣ�2:17   2004-12-14   
  '=================   ����˵��   ================   
  'INI.OpenFile   =   �ļ�·����ʹ������·�������ⲿ���壩   
  'INI.CodeSet     =   �������ã�Ĭ��Ϊ   GB2312   
  'INI.IsTrue       =   ����ļ��Ƿ����������ڣ�   
  '================   ����˵��   =================   
  'IsGroup(����)                         ������Ƿ����   
  'IsNode(����,�ڵ���)                         ���ڵ��Ƿ����   
  'GetGroup(����)                         ��ȡ����Ϣ   
  'CountGroup()                         ͳ��������   
  'ReadNode(����,�ڵ���)                         ��ȡ�ڵ�����   
  'WriteGroup(����)                         ������   
  'WriteNode(��,�ڵ�,�ڵ�����)             ����/���½ڵ�����   
  'DeleteGroup(����)                         ɾ����   
  'DeleteNode(����,�ڵ���)             ɾ���ڵ�   
  'Save()                                     �����ļ�   
  'Close()                                     ����ڲ����ݣ��ͷţ�   
  '===============================================   
    
    
  Class   INI_Class   
  '===============================================   
            Private   Stream                         '//   Stream   ����   
            Private   FilePath             '//   �ļ�·��   
            Public   Content                         '//   �ļ�����   
            Public   IsTrue                         '//   �ļ��Ƿ����   
            Public   IsAnsi                         '//   ��¼�Ƿ������   
            Public   CodeSet                         '//   ���ݱ���   
  '================================================   
              
            '//   ��ʼ��   
            Private   Sub   Class_Initialize()   
                        Set   Stream             =   Server.CreateObject("ADODB.Stream")   
                        Stream.Mode             =   3   
                        Stream.Type             =   2   
                        CodeSet                         =   "gb2312"   
                        IsAnsi                         =   True   
                        IsTrue                         =   True   
            End   Sub   
              
              
            '//   ��������ת��Ϊ�ַ���   
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
              
              
            '//   �����ļ�·��   
            Property   Let   OpenFile(INIFilePath)   
                        FilePath   =   INIFilePath   
                        Stream.Open   
                        On   Error   Resume   Next   
                        Stream.LoadFromFile(FilePath)   
                        '//   �ļ�������ʱ���ظ�   IsTrue   
                        if   Err.Number<>0   Then   
                                    IsTrue   =   False   
                                    Err.Clear   
                        End   if   
                        Content   =   Stream.ReadText(Stream.Size)   
                        if   Not   IsAnsi   Then   Content=Bytes2bStr(Content)   
            End   Property   
              
              
            '//   ������Ƿ����[����:����]   
            Public   Function   IsGroup(GroupName)   
                        if   Instr(Content,"["&GroupName&"]")>0   Then   
                                    IsGroup   =   True   
                        Else   
                                    IsGroup   =   False   
                        End   if   
            End   Function   
              
              
            '//   ��ȡ����Ϣ[����:����]   
            Public   Function   GetGroup(GroupName)   
                        Dim   TempGroup   
                        if   Not   IsGroup(GroupName)   Then   Exit   Function   
                        '//   ��ʼѰ��ͷ����ȡ   
                        TempGroup   =   Mid(Content,Instr(Content,"["&GroupName&"]"),Len(Content))   
                        '//   �޳�β��   
                        if   Instr(TempGroup,VbCrlf&"[")>0   Then   TempGroup=Left(TempGroup,Instr(TempGroup,VbCrlf&"[")-1)   
                        if   Right(TempGroup,1)<>Chr(10)   Then   TempGroup=TempGroup&VbCrlf   
                        GetGroup   =   TempGroup   
            End   Function   
              
              
            '//   ���ڵ��Ƿ����[����:����,�ڵ���]   
            Public   Function   IsNode(GroupName,NodeName)   
                        if   Instr(GetGroup(GroupName),NodeName&"=")   Then   
                                    IsNode   =   True   
                        Else   
                                    IsNode   =   False   
                        End   if   
            End   Function   
              
              
            '//   ������[����:����]   
            Public   Sub   WriteGroup(GroupName)   
                        if   Not   IsGroup(GroupName)   And   GroupName<>""   Then   
                                    Content   =   Content   &   "["   &   GroupName   &   "]"   &   VbCrlf   
                        End   if   
            End   Sub   
              
              
            '//   ��ȡ�ڵ�����[����:����,�ڵ���]   
            Public   Function   ReadNode(GroupName,NodeName)   
                        if   Not   IsNode(GroupName,NodeName)   Then   Exit   Function   
                        Dim   TempContent   
                        '//   ȡ����Ϣ   
                        TempContent   =   GetGroup(GroupName)   
                        '//   ȡ��ǰ�ڵ�����   
                        TempContent   =   Right(TempContent,Len(TempContent)-Instr(TempContent,NodeName&"=")+1)   
                        TempContent   =   Replace(Left(TempContent,Instr(TempContent,VbCrlf)-1),NodeName&"=","")   
                        ReadNode   =   ReplaceData(TempContent,0)   
            End   Function   
              
              
            '//   д��ڵ�����[����:����,�ڵ���,�ڵ�����]   
            Public   Sub   WriteNode(GroupName,NodeName,NodeData)   
                        '//   �鲻����ʱд����   
                        if   Not   IsGroup(GroupName)   Then   WriteGroup(GroupName)   
                          
                        '//   Ѱ��λ�ò�������   
                        '///   ��ȡ��   
                        Dim   TempGroup   :   TempGroup   =   GetGroup(GroupName)   
                          
                        '///   ����β��׷��   
                        Dim   NewGroup   
                        if   IsNode(GroupName,NodeName)   Then   
                                    NewGroup   =   Replace(TempGroup,NodeName&"="&ReplaceData(ReadNode(GroupName,NodeName),1),NodeName&"="&ReplaceData(NodeData,1))   
                        Else   
                                    NewGroup   =   TempGroup   &   NodeName   &   "="   &   ReplaceData(NodeData,1)   &   VbCrlf   
                        End   if   
                          
                        Content   =   Replace(Content,TempGroup,NewGroup)   
            End   Sub   
              
              
            '//   ɾ����[����:����]   
            Public   Sub   DeleteGroup(GroupName)   
                        Content   =   Replace(Content,GetGroup(GroupName),"")   
            End   Sub   
              
              
            '//   ɾ���ڵ�[����:����,�ڵ���]   
            Public   Sub   DeleteNode(GroupName,NodeName)   
                        Dim   TempGroup   
                        Dim   NewGroup   
                        TempGroup   =   GetGroup(GroupName)   
                        NewGroup   =   Replace(TempGroup,NodeName&"="&ReadNode(GroupName,NodeName)&VbCrlf,"")   
                        if   Right(NewGroup,1)<>Chr(10)   Then   NewGroup   =   NewGroup&VbCrlf   
                        Content   =   Replace(Content,TempGroup,NewGroup)   
            End   Sub   
              
              
            '//   �滻�ַ�[ʵ��:�滻Ŀ��,����������]   
            '             �ַ�ת��[��ֹ�ؼ����ų���]   
            '             [                         --->             {(@)}   
            '             ]                         --->             {(#)}   
            '             =                         --->             {($)}   
            '             �س�             --->             {(1310)}   
            Public   Function   ReplaceData(Data_Str,IsIn)   
                        if   IsIn   Then   
                                    ReplaceData   =   Replace(Replace(Replace(Data_Str,"[","{(@)}"),"]","{(#)}"),"=","{($)}")   
                                    ReplaceData   =   Replace(ReplaceData,Chr(13)&Chr(10),"{(1310)}")   
                        Else   
                                    ReplaceData   =   Replace(Replace(Replace(Data_Str,"{(@)}","["),"{(#)}","]"),"{($)}","=")   
                                    ReplaceData   =   Replace(ReplaceData,"{(1310)}",Chr(13)&Chr(10))   
                        End   if   
            End   Function   
              
              
            '//   �����ļ�����   
            Public   Sub   Save()   
                        With   Stream   
                                    .Close   
                                    .Open   
                                    .WriteText   Content   
                                    .SaveToFile   FilePath,2   
                        End   With   
            End   Sub   
              
              
            '//   �رա��ͷ�   
            Public   Sub   Close()   
                        Set   Stream   =   Nothing   
                        Set   Content   =   Nothing   
            End   Sub   
              
  End   Class   
    
    
  Dim   INI   
  Set   INI   =   New   INI_Class   
  INI.OpenFile   =   Server.MapPath("Config.ini")   
  '==========   ����д��ini����   ==========   
  INI.WriteNode("SiteConfig","SiteName","Leadbbs������̳")   
  INI.WriteNode("SiteConfig","Mail","leadbbs@leadbbs.com")   
  INI.Save()   
  '==========   ���Ƕ�ȡini����   ==========   
  Response.Write("վ�����ƣ�"INI.ReadNode("SiteConfig","SiteName"))   
  %> 