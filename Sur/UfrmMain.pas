unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,StrUtils, DB,ComObj,Variants,
  ScktComp,EncdDecd{DecodeStream},Jpeg{TJPEGImage}, IdBaseComponent, IdCoder,
  IdCoder3to4, IdCoderMIME;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
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
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
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
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{�����ļ���Ч}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
    function DIFF_decode(const ASTMField:string):string;
    //function GetSpecNo(const Value:string):string; //ȡ��������
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//�ӽ�������
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='����!���뿪������ϵ!' ;
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
  ifRecLog:boolean;//�Ƿ��¼������־
  NoDtlStr:integer;//������ʶλ
  ifSocketClient:boolean;
  ifKLite8:boolean;
  KLite8_Patient_ID:boolean;

  RFM:STRING;       //��������
  hnd:integer;
  bRegister:boolean;

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

  if not result then messagedlg('�Բ���,��û��ע���ע�������,��ע��!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//�Ƿ񼯳ɵ�¼ģʽ

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('�������ݿ�', '������', '');
  initialcatalog := Ini.ReadString('�������ݿ�', '���ݿ�', '');
  ifIntegrated:=ini.ReadBool('�������ݿ�','���ɵ�¼ģʽ',false);
  userid := Ini.ReadString('�������ݿ�', '�û�', '');
  password := Ini.ReadString('�������ݿ�', '����', '107DFC967CDCFAAF');
  Ini.Free;
  //======����password
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
  //Persist Security Info,��ʾADO�����ݿ����ӳɹ����Ƿ񱣴�������Ϣ
  //ADOȱʡΪTrue,ADO.netȱʡΪFalse
  //�����лᴫADOConnection��Ϣ��TADOLYQuery,������ΪTrue
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  rfm:='';

  ConnectString:=GetConnectString;
  
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  lytray1.Hint:='���ݽ��շ���'+ExtractFileName(Application.ExeName);

//=============================��ʼ������=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('��л��ʹ�����ܼ��ϵͳ��'+chr(13)+'���ס��ʼ�����룺'+'lc'),
        //            'ϵͳ��ʾ',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    if LoadInputPassDll then begin ServerSocket1.Close;action:=cafree;end else action:=caNone;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
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

  ifSocketClient:=ini.readBool(IniSection,'Socket�ͻ���',false);//BC-10:�ͻ���
  ServerIP:=trim(ini.ReadString(IniSection,'������IP',''));
  ServerPort:=ini.ReadInteger(IniSection,'�������˿�',8080);//DH36��Ĭ�϶˿�Ϊ5600
  NoDtlStr:=ini.ReadInteger(IniSection,'������ʶλ',3);//BS300:4;DH36��BC-10:3

  autorun:=ini.readBool(IniSection,'�����Զ�����',false);
  ifRecLog:=ini.readBool(IniSection,'������־',false);
  ifKLite8:=ini.readBool(IniSection,'KLite8��Ӧ',false);
  KLite8_Patient_ID:=ini.readBool(IniSection,'KLite8������',false);

  GroupName:=trim(ini.ReadString(IniSection,'������',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'������ĸ','')));//�������Ǵ�д������һʧ��
  SpecType:=ini.ReadString(IniSection,'Ĭ����������','');
  SpecStatus:=ini.ReadString(IniSection,'Ĭ������״̬','');
  CombinID:=ini.ReadString(IniSection,'�����Ŀ����','');

  LisFormCaption:=ini.ReadString(IniSection,'����ϵͳ�������','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9997');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  ServerSocket1.Close;
  ServerSocket1.Port:=ServerPort;
  ClientSocket1.Close;
  ClientSocket1.Port:=ServerPort;
  ClientSocket1.OnRead:=ServerSocket1ClientRead;//Client��Server�Ķ����ݷ���һģһ��
  if ifSocketClient then
  begin
    ClientSocket1.Host:=ServerIP;
    Timer1.Interval:=5000;
    Timer1.Enabled:=true;
    try
      ClientSocket1.Open;
    except
      showmessage('���ӷ�����'+ServerIP+'('+inttostr(ServerPort)+')ʧ��!');
    end;
  end else
  begin
    try
      ServerSocket1.Open;
    except
      showmessage('�˿�'+inttostr(ServerPort)+'��ʧ��!');
    end;
  end;
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
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
    ss:='������'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ݿ�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ɵ�¼ģʽ'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '�û�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '����'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('�������ݿ�','�������ݿ�',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='Socket�ͻ���'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '������IP'+#2+'Edit'+#2+#2+'1'+#2+'��λ��ͨ�Žӿڳ���Ϊ��������ʱ������д'+#2+#3+
      '�������˿�'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ʶλ'+#2+'Edit'+#2+#2+'1'+#2+'OBX���ô��߷ָ�,��0��ʼ,�ڼ�λ'+#2+#3+
      '������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ĸ'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '����ϵͳ�������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ����������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ������״̬'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Ŀ����'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Զ�����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '������־'+#2+'CheckListBox'+#2+#2+'0'+#2+'ע:ǿ�ҽ�������������ʱ�ر�'+#2+#3+
      'KLite8��Ӧ'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      'KLite8������'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
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
  showmessage('����ɹ�!');
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
  ServerSocket1ClientRead(nil,nil);
  ls.Free;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'���ô���������ϵ��ַ�������������,�Ի�ȡע����'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('ע��:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
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
  i,j:integer;
  Str:string;
  SBPos,EBPos:integer;
  ls,ls2,ls3,ls4,ls5:tstrings;
  DtlStr:string;
  CheckDate:string;
  sHistogramTemp:string;
  sHistogramString:string;
  sHistogramFile:string;
  strList:TStrings;
  Message_Control_ID:string;
begin
  Str:=Socket.ReceiveText;
  if length(memo1.Lines.Text)>=60000 then memo1.Lines.Clear;//memoֻ�ܽ���64K���ַ�
  memo1.Lines.Add(Str);

  rfm:=rfm+Str;
  
  SBPos:=pos(#$0B,rfm);
  if SBPos<=0 then exit;
  delete(rfm,1,SBPos-1);//����ͷ�ǵ�һ���ַ�

  EBPos:=pos(#$1C#$0D,rfm);
  while EBPos>0 do
  begin
    rfm2:=copy(rfm,1,EBPos+1);//1���걾���
    delete(rfm,1,EBPos+1);

    SpecNo:=formatdatetime('nnss',now);

    ls:=TStringList.Create;
    ExtractStrings([#$D],[],Pchar(rfm2),ls);

    ReceiveItemInfo:=VarArrayCreate([0,ls.Count-1],varVariant);

    for  i:=0  to ls.Count-1 do
    begin
      if uppercase(copy(trim(ls[i]),1,4))='MSH|' then
      begin
        ls5:=StrToList(ls[i],'|');
        if ls5.Count>9 then Message_Control_ID:=ls5[9];
        ls5.Free;
      end;
      
      if uppercase(copy(trim(ls[i]),1,4))='OBR|' then
      begin
        ls3:=StrToList(ls[i],'|');

        if ls3.Count>3 then SpecNo:=rightstr('0000'+ls3[3],4);

        if KLite8_Patient_ID and(ls3.Count>2) then
        begin
          SpecNo:=rightstr('0000'+StringReplace(ls3[2],'^R','',[rfReplaceAll, rfIgnoreCase]),4);
        end;

        if ls3.Count>7 then
          CheckDate:=copy(ls3[7],1,4)+'-'+copy(ls3[7],5,2)+'-'+copy(ls3[7],7,2)+' '+copy(ls3[7],9,2)+ifThen(copy(ls3[7],9,2)<>'',':')+copy(ls3[7],11,2);
        ls3.Free;
      end;
      
      DtlStr:='';
      sValue:='';
      sHistogramString:='';
      sHistogramFile:='';
      if uppercase(copy(trim(ls[i]),1,4))='OBX|' then
      begin
        ls2:=StrToList(ls[i],'|');
        if(ls2.Count>5)and(ls2.Count>NoDtlStr)then
        begin
          DtlStr:=ls2[NoDtlStr];
          sValue:=ls2[5];
        end;

        //ֱ��ͼ���� start DH36
        if (POS('Histogram. BMP',DtlStr)>0)and(ls2.Count>5) then
        begin
          sValue:='';
          sHistogramString:='';

          ls4:=StrToList(ls2[5],'^');//ls2[5]Ϊ^Image^PNG^Base64^iVBORw0KGgoAAAANSUhEUgAAAJw.........
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
          
          strList:=TStringlist.Create;
          try
            strList.Add(sHistogramTemp);
            strList.SaveToFile(sHistogramFile);
          finally
            strList.Free;
          end;
        end;
        //ֱ��ͼ���� stop

        //ֱ��ͼ���� start URIT-2980
        if (('WBCHistogram'=DtlStr)or('RBCHistogram'=DtlStr)or('PLTHistogram'=DtlStr))and(ls2.Count>5) then
        begin
          sValue:='';
          sHistogramFile:='';

          //����PLTͼ�����ݵ�ѡȡ��
          //2900P����3.64.xxxx�Ժ�İ汾
          //3020/3000P����4.64.xxxx�Ժ�İ汾
          //3060/3080/3081����6.65.xxxx�Ժ�İ汾
          //2960/2980/2981����5.65.xxxx�Ժ�İ汾
          //��ֻȡPLTͼ�����ݵ�ǰ100���ֽڣ�������֮ǰ�İ汾��ȡǰ128���ֽ�

          ls4:=StrToList(ls2[5],'^');//ls2[5]Ϊ^Histogram^32Byte^HEX^00000000000000000.........
          sHistogramTemp:=ls4[4];
          if 'PLTHistogram'=DtlStr then sHistogramTemp:=copy(sHistogramTemp,1,200);//Ĭ�Ϲ�128���ֽ�256���ַ�
          if ls4.Count>4 then sHistogramString:=DIFF_decode(sHistogramTemp);
          ls4.Free;
        end;
        //ֱ��ͼ���� stop
        
        ls2.Free;
      end;
      ReceiveItemInfo[i]:=VarArrayof([DtlStr,sValue,sHistogramString,sHistogramFile]);

      //�����������Start
      //DH36Ӧ�ò���Ҫ��������������������ҲûӰ��
      for  j:=0  to i-1 do
      begin
        if (DtlStr<>'')and(ReceiveItemInfo[j][0]=DtlStr) then ReceiveItemInfo[j]:=VarArrayof(['','','','']);
      end;
      //�����������End
    end;
    ls.Free;

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,(SpecNo),CheckDate,
        (GroupName),(SpecType),(SpecStatus),(EquipChar),
        (CombinID),'',(LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        ifRecLog,true,'����');
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
    end;

    EBPos:=pos(#$1C#$0D,rfm);
    
    if ifKLite8 then
    begin
      Socket.SendText(#$0B+'MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||ASCII|||'+#$0D+'MSA|AA|'+Message_Control_ID+'|message accepted|||0|'+#$0D#$1C#$0D);
    end;
    
  end;
end;

procedure TfrmMain.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('�ͻ���'+Socket.RemoteHost+'('+Socket.RemoteAddress+')�Ѿ�����');
end;

procedure TfrmMain.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('�ͻ���'+Socket.RemoteHost+'('+Socket.RemoteAddress+')�Ѿ��Ͽ�');
end;

procedure TfrmMain.ServerSocket1ClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
  Memo1.Lines.Add('�ͻ���'+Socket.RemoteHost+'('+Socket.RemoteAddress+')��������');
  ErrorCode := 0;
end;

procedure TfrmMain.ServerSocket1GetSocket(Sender: TObject; Socket: Integer;
  var ClientSocket: TServerClientWinSocket);
begin
  //Memo1.Lines.Add('�ͻ�����������...');
end;

procedure TfrmMain.ServerSocket1Listen(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  //Memo1.Lines.Add('�ȴ��ͻ�������...');
end;

procedure TfrmMain.ClientSocket1Connect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('�Ѿ����ӵ�'+Socket.RemoteHost+'('+Socket.RemoteAddress+')');
end;

procedure TfrmMain.ClientSocket1Disconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
  Memo1.Lines.Add('�Ѿ��Ͽ���'+Socket.RemoteHost+'('+Socket.RemoteAddress+')������');
end;

procedure TfrmMain.ClientSocket1Error(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
  Memo1.Lines.Add('���������'+Socket.RemoteHost+'('+Socket.RemoteAddress+')�����ӷ�������');
  ErrorCode := 0;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
  if not ifSocketClient then exit;
  if ClientSocket1.Active then exit;

  try
    ClientSocket1.Open;
  except
    showmessage('���ӷ�����ʧ��!');
  end;
end;

function TfrmMain.DIFF_decode(const ASTMField: string): string;
var
  sList:TStrings;
  ss:string;
  i:integer;
begin
  ss:=ASTMField;
  
  sList:=TStringList.Create;
  while length(ss)>=2 do
  begin
    sList.Add(copy(ss,1,2));
    delete(ss,1,2);
  end;
  for i :=0  to sList.Count-1 do
  begin
    result:=result+' '+inttostr(strtoint('$'+sList[i]));
  end;
  sList.Free;
  result:=trim(result);
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('�ó������������У�'),
                    'ϵͳ��ʾ',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
