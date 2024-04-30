unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  Menus, StdCtrls, Buttons, ADODB,
  ComCtrls, ToolWin, ExtCtrls,
  inifiles,Dialogs,StrUtils, DB,ComObj,Variants,
  ScktComp,EncdDecd{DecodeStream},Jpeg{TJPEGImage}, IdBaseComponent, IdCoder,
  IdCoder3to4, IdCoderMIME, CoolTrayIcon, Uni,OracleUniProvider;

type
  TfrmMain = class(TForm)
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button1: TButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ServerSocket1: TServerSocket;
    SaveDialog1: TSaveDialog;
    IdDecoderMIME1: TIdDecoderMIME;
    ClientSocket1: TClientSocket;
    Timer1: TTimer;
    LYTray1: TCoolTrayIcon;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ServerSocket1ClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientError(Sender: TObject;
      Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
      var ErrorCode: Integer);
    procedure ServerSocket1GetSocket(Sender: TObject; Socket: Integer;
      var ClientSocket: TServerClientWinSocket);
    procedure ServerSocket1Listen(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Connect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Disconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Error(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    procedure UpdateConfig;{配置文件生效}
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

type TMsgType = (QRY, ORU);//QRY表示仪器向LIS查询病人信息;ORU表示仪器向LIS发送结果

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  ifRecLog:boolean;//是否记录调试日志
  EquipUnid:integer;//设备唯一编号
  NoDtlStr:integer;//联机标识位
  ifSocketClient:boolean;
  ifKLite8:boolean;
  Line_Patient_ID:String;
  No_Patient_ID:integer;
  FS205_Chinese:boolean;
  BS300_Rerun:boolean;
  Discard_Qualitative:boolean;//丢弃逗号分隔的定性结果
  HisConnStr:String;
  CM_Category_Message:String;//响应消息类型

  RFM:STRING;       //返回数据
  hnd:integer;
  bRegister:boolean;

  if_test:boolean;//是否测试

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  if_test:=false;
  rfm:='';

  ConnectString:=GetConnectString;
  
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  Caption:='数据接收服务'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=caNone;
  LYTray1.HideMainForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  if (MessageDlg('退出后将不再接收设备数据,确定退出吗？', mtWarning, [mbYes, mbNo], 0) <> mrYes) then exit;
  application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  LYTray1.ShowMainForm;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
  ServerPort:integer;
  ServerIP:string;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  ifSocketClient:=ini.readBool(IniSection,'Socket客户端',false);//BC-10:客户端
  ServerIP:=trim(ini.ReadString(IniSection,'服务器IP',''));
  ServerPort:=ini.ReadInteger(IniSection,'服务器端口',8080);//DH36的默认端口为5600
  NoDtlStr:=ini.ReadInteger(IniSection,'联机标识位',3);//BS300:4;DH36、BC-10:3

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);
  EquipUnid:=ini.ReadInteger(IniSection,'设备唯一编号',-1);
  ifKLite8:=ini.readBool(IniSection,'KLite8响应',false);
  Line_Patient_ID:=ini.ReadString(IniSection,'联机号所在行','');
  No_Patient_ID:=ini.ReadInteger(IniSection,'联机号位',3);
  CM_Category_Message:=ini.ReadString(IniSection,'响应消息类型','');

  FS205_Chinese:=ini.readBool(IniSection,'中文乱码解码',false);
  BS300_Rerun:=ini.readBool(IniSection,'处理BS300重做',false);
  Discard_Qualitative:=ini.readBool(IniSection,'丢弃逗号分隔的定性结果',false);

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  HisConnStr:=ini.ReadString(IniSection,'连接HIS数据库','');
  
  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  ServerSocket1.Close;
  ServerSocket1.Port:=ServerPort;
  ClientSocket1.Close;
  ClientSocket1.Port:=ServerPort;
  ClientSocket1.OnRead:=ServerSocket1ClientRead;//Client与Server的读数据方法一模一样
  if ifSocketClient then
  begin
    ClientSocket1.Host:=ServerIP;
    Timer1.Interval:=5000;
    Timer1.Enabled:=true;
    try
      ClientSocket1.Open;
    except
      showmessage('连接服务器'+ServerIP+'('+inttostr(ServerPort)+')失败!');
    end;
  end else
  begin
    try
      ServerSocket1.Open;
    except
      showmessage('端口'+inttostr(ServerPort)+'打开失败!');
    end;
  end;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  ss:='Socket客户端'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '服务器IP'+#2+'Edit'+#2+#2+'1'+#2+'上位机通信接口程序为服务器端时无需填写'+#2+#3+
      '服务器端口'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '联机标识位'+#2+'Edit'+#2+#2+'1'+#2+'OBX行用垂线分隔,从0开始,第几位'+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Combobox'+#2+'<自行填写>'+#13+'OBR第15位'+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '设备唯一编号'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'KLite8响应'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '响应消息类型'+#2+'Combobox'+#2+'ACK^R01'+#13+'ACK'+#2+'1'+#2+#2+#3+
      '联机号所在行'+#2+'Combobox'+#2+'PID'+#13+'OBR'+#2+'1'+#2+#2+#3+
      '联机号位'+#2+'Edit'+#2+#2+'1'+#2+'PID或OBR行用垂线分隔,从0开始,第几位'+#2+#3+
      '中文乱码解码'+#2+'CheckListBox'+#2+#2+'1'+#2+'判断依据:中文及特殊字符(如μ)是否显示正常'+#2+#3+
      '处理BS300重做'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '丢弃逗号分隔的定性结果'+#2+'CheckListBox'+#2+#2+'1'+#2+'iFlash-3000结果【2.05,无反应性】'+#2+#3+
      '连接HIS数据库'+#2+'UniConn'+#2+#2+'1'+#2+'Oracle Server格式:IP:Port:SID'+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Memo1.Lines.Clear;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  SaveDialog1.DefaultExt := '.txt';
  SaveDialog1.Filter := 'txt (*.txt)|*.txt';
  if not SaveDialog1.Execute then exit;
  memo1.Lines.SaveToFile(SaveDialog1.FileName);
  showmessage('保存成功!');
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
  ls:Tstrings;
begin
  OpenDialog1.DefaultExt := '.txt';
  OpenDialog1.Filter := 'txt (*.txt)|*.txt';
  if not OpenDialog1.Execute then exit;
  ls:=Tstringlist.Create;
  ls.LoadFromFile(OpenDialog1.FileName);
  rfm:=ls.Text;
  if_test:=true;
  ServerSocket1ClientRead(nil,nil);
  if_test:=false;
  ls.Free;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.ServerSocket1ClientRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
  SpecNo:string;
  rfm2:string;
  sValue:string;
  FInts:OleVariant;
  ReceiveItemInfo:OleVariant;
  i,j{,k}:integer;
  Str:string;
  SBPos,EBPos:integer;
  ls,ls2,ls3,ls4,ls5,ls7,ls8,ls9:tstrings;
  DtlStr:string;
  CheckDate:string;
  sHistogramTemp:string;
  sHistogramFile:string;
  strList:TStrings;
  Message_Control_ID:string;
  //Query_Target:String;
  //ORF:String;

  //lls:TStrings;
  //UniConnection1:TUniConnection;
  //UniQuery1:TUniQuery;

  //patientname:String;
  //sex:String;
  //age:String;
  //AGEUNIT:String;
  //report_date:String;//申请时间
  //deptname:String;//申请科室
  //check_doctor:String;//申请医生
  //His_Unid:String;
  //s1:String;
  //His_ItemId:String;
  //ItemList: TStrings;

  //Conn:TADOConnection;
  //AdoQry:TAdoQuery;
  
  //MsgType:TMsgType;

  r_Barcode:String;
  r_patientname:String;
  r_sex:String;
  r_age:String;
  r_check_doctor:String;//申请医生
begin
  if not if_test then Str:=Socket.ReceiveText;
  if FS205_Chinese then Str:=UTF8Decode(Str);//解决【飞测FS-205】中文乱码问题
  
  if length(memo1.Lines.Text)>=1000000 then memo1.Lines.Clear;//memo在win98只能接受64K个字符,在win2000无限制
  memo1.Lines.Add(Str);

  rfm:=rfm+Str;
  
  SBPos:=pos(#$0B,rfm);
  if SBPos<=0 then exit;
  delete(rfm,1,SBPos-1);//保持头是第一个字符

  EBPos:=pos(#$1C#$0D,rfm);
  while EBPos>0 do
  begin
    rfm2:=copy(rfm,1,EBPos+1);//1个标本结果
    delete(rfm,1,EBPos+1);

    EBPos:=pos(#$1C#$0D,rfm);

    //if (pos(#$0D'QRD|',rfm2)>0)and(pos(#$0D'QRF|',rfm2)>0) then MsgType:=QRY else MsgType:=ORU;

    //case MsgType of
      {QRY: begin
        if trim(HisConnStr)='' then continue;
        
        ls:=TStringList.Create;
        ExtractStrings([#$D],[],Pchar(rfm2),ls);

        for  i:=0  to ls.Count-1 do
        begin
          if uppercase(copy(trim(ls[i]),1,4))='MSH|' then
          begin
            ls5:=StrToList(ls[i],'|');
            if ls5.Count>9 then Message_Control_ID:=ls5[9];
            ls5.Free;
          end;

          if uppercase(copy(trim(ls[i]),1,4))='QRD|' then
          begin
            ls5:=StrToList(ls[i],'|');
            if ls5.Count>8 then Query_Target:=ls5[8];
            ls5.Free;
          end;        
        end;

        ls.Free;

        //读HIS数据库
        lls:=TStringList.Create;
        lls.Delimiter:=';';
        lls.DelimitedText:=HisConnStr;
      
        UniConnection1:=TUniConnection.Create(nil);
        UniConnection1.ProviderName:=lls.Values['ProviderName'];
        UniConnection1.SpecificOptions.Values['Direct']:='True';
        UniConnection1.LoginPrompt:=false;
        UniConnection1.Username:=lls.Values['Username'];
        UniConnection1.Password:=lls.Values['Password'];
        UniConnection1.Server:=lls.Values['Server'];

        lls.Free;

        Try
          UniConnection1.Connect;
        except
          on E:Exception do
          begin
            memo1.Lines.Add('连接HIS数据库失败:'+E.Message);
          end;
        end;

        if not UniConnection1.Connected then begin UniConnection1.Free;continue;end;

        UniQuery1:=TUniQuery.Create(nil);
        UniQuery1.Connection:=UniConnection1;
        UniQuery1.SQL.Text:='select wm_concat(distinct ORDER_ID) as ORDER_ID from LIS_REQUEST where BARCODE='''+copy(Query_Target,Pos('^',Query_Target)+1,MaxInt)+''' ';//用逗号连接HIS组合项目代码
        Try
          UniQuery1.Open;
        except
          on E:Exception do
          begin
            memo1.Lines.Add('查询HIS组合项目失败:'+E.Message);
          end;
        end;
        
        if not UniQuery1.Active then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        if UniQuery1.RecordCount<=0 then begin UniQuery1.Free;UniConnection1.Free;continue;end;

        aJson:=SO(GetLisCombItem(PChar(ConnectString),PChar(UniQuery1.fieldbyname('ORDER_ID').AsString),PChar(EquipChar),'PEIS'));

        if not aJson.AsObject.Exists('项目信息') then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        aSuperArray:=aJson['项目信息'].AsArray;
  
        ItemList:=TStringList.Create;
        for k:=0 to aSuperArray.Length-1 do
        begin
          ItemList.Add(aSuperArray[k]['组合项目备注'].AsString);//利用组合项目的【说明】字段，1-全部、0-沉渣、2-干化学
        end;

        //组合项目,1-全部、0-沉渣、2-干化学
        if (ItemList.IndexOf('1')>=0) or ((ItemList.IndexOf('0')>=0) and (ItemList.IndexOf('2')>=0)) then s1:='1'
          else if ItemList.IndexOf('0')>=0 then s1:='0'
            else if ItemList.IndexOf('2')>=0 then s1:='2'
              else begin ItemList.Free;UniQuery1.Free;UniConnection1.Free;continue;end;

        ItemList.Free;

        UniQuery1.Close;
        UniQuery1.SQL.Clear;
        UniQuery1.SQL.Text:='select top 1 * from LIS_REQUEST where BARCODE='''+copy(Query_Target,Pos('^',Query_Target)+1,MaxInt)+''' ';//用逗号连接HIS组合项目代码
        Try
          UniQuery1.Open;
        except
          on E:Exception do
          begin
            memo1.Lines.Add('查询HIS受检者信息失败:'+E.Message);
          end;
        end;

        if not UniQuery1.Active then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        if UniQuery1.RecordCount<=0 then begin UniQuery1.Free;UniConnection1.Free;continue;end;

        patientname:=UniQuery1.fieldbyname('NAME').AsString;//患者姓名
        //患者性别
        if '1'=UniQuery1.fieldbyname('SEX').AsString then sex:='男'
          else if '2'=UniQuery1.fieldbyname('SEX').AsString then sex:='女'
            else if '3'=UniQuery1.fieldbyname('SEX').AsString then sex:='不详';
        if UniQuery1.fieldbyname('AGEUNIT').AsString='Y' THEN AGEUNIT:='岁'
          else if UniQuery1.fieldbyname('AGEUNIT').AsString='N' THEN AGEUNIT:='月';
        age:=UniQuery1.fieldbyname('AGE').AsString+AGEUNIT;//患者年龄
        report_date:=UniQuery1.fieldbyname('WRITE_TIME').AsString;//申请时间
        deptname:=UniQuery1.fieldbyname('REQDEPT').AsString;//申请科室
        check_doctor:=UniQuery1.fieldbyname('WRITE_NAME').AsString;//申请医生
        His_Unid:=UniQuery1.fieldbyname('REG_ID').AsString;//体检号

        UniQuery1.Free;
        UniConnection1.Free;

        ORF:=#$0B+
             'MSH|^~\&|||||||ORF|1|P|2.3'+#$0D+
             'MSA|AA|'+Message_Control_ID+#$0D+
             'QRD||R|I||||20^LI|'+Query_Target+'|DEM|ALL'+#$0D+
             'PID|||'+Query_Target+'||'+s1+'|'+patientname+'|'+His_Unid+'|'+age+'|'+sex+#$0D+//PID-3此域包含用于标识患者身份的唯一标识号。这里是样本号和条码。是不是返回结果的样本号？为空会怎么样？//PID-5测试模式“1”or”0”or”2”(全部、沉渣、干化学)
             'PV1|||'+#$0D+
             'OBR|||||||||||||||'+deptname+'|'+check_doctor+#$0D+
             #$1C#$0D;
        Socket.SendText(ORF);
      end;//}
      //ORU: begin
        ls:=TStringList.Create;
        ExtractStrings([#$D],[],Pchar(rfm2),ls);

        SpecNo:='';

        r_Barcode:='';
        r_patientname:='';
        r_sex:='';
        r_age:='';
        r_check_doctor:='';//申请医生

        ReceiveItemInfo:=VarArrayCreate([0,ls.Count-1],varVariant);

        for  i:=0  to ls.Count-1 do
        begin
          if uppercase(copy(trim(ls[i]),1,4))='MSH|' then
          begin
            ls5:=StrToList(ls[i],'|');
            if ls5.Count>9 then Message_Control_ID:=ls5[9];
            ls5.Free;
          end;

          if uppercase(copy(trim(ls[i]),1,4))='PID|' then
          begin
            if Line_Patient_ID='PID' then SpecNo:=ls[i];
            ls9:=StrToList(ls[i],'|');
            if ls9.Count>4 then r_Barcode:=ls9[4];
            if ls9.Count>5 then r_patientname:=ls9[5];
            if ls9.Count>8 then r_sex:=ls9[8];
            if ls9.Count>7 then r_age:=ls9[7];
            ls9.Free;
          end;

          if uppercase(copy(trim(ls[i]),1,4))='OBR|' then
          begin
            if Line_Patient_ID='OBR' then SpecNo:=ls[i];

            ls3:=StrToList(ls[i],'|');
            if ls3.Count>7 then CheckDate:=copy(ls3[7],1,4)+'-'+copy(ls3[7],5,2)+'-'+copy(ls3[7],7,2)+' '+copy(ls3[7],9,2)+ifThen(copy(ls3[7],9,2)<>'',':')+copy(ls3[7],11,2);
            if(SpecType='OBR第15位')and(ls3.Count>15) then SpecType:=ls3[15];
            if ls3.Count>16 then r_check_doctor:=ls3[16];//申请医生
            ls3.Free;
          end;

          DtlStr:='';
          sValue:='';
          sHistogramFile:='';
          if uppercase(copy(trim(ls[i]),1,4))='OBX|' then
          begin
            ls2:=StrToList(ls[i],'|');

            if ls2.Count>NoDtlStr then DtlStr:=ls2[NoDtlStr];

            if(ls2.Count>5)and(ls2[2]<>'ED')then//ls2[2]='ED'表示图片结果
            begin
              sValue:=ls2[5];
              sValue:=StringReplace(sValue,'↑','',[rfReplaceAll, rfIgnoreCase]);//飞测FS-205
              sValue:=StringReplace(sValue,'↓','',[rfReplaceAll, rfIgnoreCase]);//飞测FS-205
              sValue:=StringReplace(sValue,'mg/L','',[rfReplaceAll, rfIgnoreCase]);//EU-5300
              sValue:=StringReplace(sValue,'mmol/L','',[rfReplaceAll, rfIgnoreCase]);//EU-5300

              //FUS2000
              ls7:=StrToList(sValue,'^');
              if ls7.Count>2 then sValue:=trim(ls7[1]+' '+ls7[2]);
              ls7.Free;

              //iFlash-3000.例如,结果为【2.05,无反应性】,只取逗号前的2.05
              if Discard_Qualitative and(pos(',',sValue)>0) then sValue:=copy(sValue,1,pos(',',sValue)-1);

              //迈瑞EU-5300 begin
              if(pos(':\E\',sValue)>0)and(rightstr(sValue,4)='.JPG')then//表示结果是图片路径
              begin
                sHistogramFile:=StringReplace(sValue,'\E\','\',[rfReplaceAll]);
                sValue:='';
              end;
              //迈瑞EU-5300 end
            end;

            //图片处理 strat
            if(ls2[2]='ED')and(ls2.Count>5)and(trim(ls2[5])<>'') then//ls2[2]='ED'表示图片内容,ls2[5]表示图片内容
            begin
              //DH36:ls2[5]为^Image^PNG^Base64^iVBORw0KGgoAAAANSUhEUgAAAJw.........
              //BC10:ls2[5]为^Image^BMP^Base64^Qk0GcAAAAAAAAL.........
              ls4:=StrToList(ls2[5],'^');
              if ls4.Count>4 then
              begin
                sHistogramFile:=DtlStr+'.'+ls4[2];

                try
                  sHistogramTemp:=IdDecoderMIME1.DecodeString(ls4[4]);
                except
                  sHistogramFile:='';
                end;
              end;
              ls4.Free;

              if pos('^',ls2[5])<=0 then//FUS-2000、GMD-S600的ls2[5]都是图片数据,没有^
              begin
                sHistogramFile:=DtlStr+ifThen(leftstr(ls2[5],4)='/9j/','.jpg','.bmp');//参见文档【MUS系列全自动尿液分析系统接口规范20220210.pdf】

                try
                  //FUS2000的ls2[5]实际上可能包含多张图片,用424D分隔,424D在sHistogramTemp中
                  //但LIS不支持单个项目多张图片的保存与显示,刚好下面的处理方式也只会识别一张图片,故懒得拆分处理了
                  sHistogramTemp:=IdDecoderMIME1.DecodeString(ls2[5]);
                except
                  sHistogramFile:='';
                end;
              end;

              strList:=TStringlist.Create;
              try
                strList.Add(sHistogramTemp);
                strList.SaveToFile(sHistogramFile);
              finally
                strList.Free;
              end;
            end;
            //图片处理 stop

            ls2.Free;
          end;
          ReceiveItemInfo[i]:=VarArrayof([DtlStr,sValue,'',sHistogramFile]);

          //处理重做结果Start
          if BS300_Rerun then
          begin
            for  j:=0  to i-1 do
            begin
              if (DtlStr<>'')and(ReceiveItemInfo[j][0]=DtlStr) then ReceiveItemInfo[j]:=VarArrayof(['','','','']);
            end;
          end;
          //处理重做结果End
        end;
        
        ls.Free;

        //联机号begin
        ls8:=StrToList(SpecNo,'|');
        if ls8.Count>No_Patient_ID then SpecNo:=ls8[No_Patient_ID];
        ls8.Free;
        SpecNo:=trim(StringReplace(SpecNo,'^R','',[rfReplaceAll, rfIgnoreCase]));//KLite8
        if SpecNo='' then SpecNo:=formatdatetime('nnss',now);
        SpecNo:=rightstr('0000'+SpecNo,4);
        //联机号end

        if bRegister then
        begin
          FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
          FInts.fData2Lis(ReceiveItemInfo,(SpecNo),CheckDate,
            (GroupName),(SpecType),(SpecStatus),(EquipChar),
            (CombinID),r_patientname+'{!@#}'+r_sex+'{!@#}{!@#}'+r_age+'{!@#}{!@#}{!@#}'+r_check_doctor,(LisFormCaption),(ConnectString),
            (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
            ifRecLog,true,'常规',
            r_Barcode,
            EquipUnid,
            '','','','',
            -1,-1,-1,-1,
            -1,-1,-1,-1,
            false,false,false,false);
          if not VarIsEmpty(FInts) then FInts:= unAssigned;
        end;

        if ifKLite8 then
        begin
          Socket.SendText(#$0B+'MSH|^~$&|||||||'+CM_Category_Message+'|1|P|2.4||||0||ASCII|||'+#$0D+'MSA|AA|'+Message_Control_ID+'|message accepted|||0|'+#$0D#$1C#$0D);
        end;
      //end;
    //end;
  end;
end;

procedure TfrmMain.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('客户端'+Socket.RemoteHost+'('+Socket.RemoteAddress+')已经连接');
end;

procedure TfrmMain.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('客户端'+Socket.RemoteHost+'('+Socket.RemoteAddress+')已经断开');
end;

procedure TfrmMain.ServerSocket1ClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
  Memo1.Lines.Add('客户端'+Socket.RemoteHost+'('+Socket.RemoteAddress+')发生错误');
  ErrorCode := 0;
end;

procedure TfrmMain.ServerSocket1GetSocket(Sender: TObject; Socket: Integer;
  var ClientSocket: TServerClientWinSocket);
begin
  //Memo1.Lines.Add('客户端正在连接...');
end;

procedure TfrmMain.ServerSocket1Listen(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  //Memo1.Lines.Add('等待客户端连接...');
end;

procedure TfrmMain.ClientSocket1Connect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('已经连接到'+Socket.RemoteHost+'('+Socket.RemoteAddress+')');
end;

procedure TfrmMain.ClientSocket1Disconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('已经断开与'+Socket.RemoteHost+'('+Socket.RemoteAddress+')的连接');
end;

procedure TfrmMain.ClientSocket1Error(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
  Memo1.Lines.Add('与服务器端'+Socket.RemoteHost+'('+Socket.RemoteAddress+')的连接发生错误');
  ErrorCode := 0;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
  if not ifSocketClient then exit;
  if ClientSocket1.Active then exit;

  try
    ClientSocket1.Open;
  except
    showmessage('连接服务器失败!');
  end;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
