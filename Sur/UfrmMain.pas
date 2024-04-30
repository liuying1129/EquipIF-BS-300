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
    procedure UpdateConfig;{�����ļ���Ч}
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

type TMsgType = (QRY, ORU);//QRY��ʾ������LIS��ѯ������Ϣ;ORU��ʾ������LIS���ͽ��

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
  EquipUnid:integer;//�豸Ψһ���
  NoDtlStr:integer;//������ʶλ
  ifSocketClient:boolean;
  ifKLite8:boolean;
  Line_Patient_ID:String;
  No_Patient_ID:integer;
  FS205_Chinese:boolean;
  BS300_Rerun:boolean;
  Discard_Qualitative:boolean;//�������ŷָ��Ķ��Խ��
  HisConnStr:String;
  CM_Category_Message:String;//��Ӧ��Ϣ����

  RFM:STRING;       //��������
  hnd:integer;
  bRegister:boolean;

  if_test:boolean;//�Ƿ����

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
begin
  if_test:=false;
  rfm:='';

  ConnectString:=GetConnectString;
  
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  Caption:='���ݽ��շ���'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='���ݽ��շ���'+ExtractFileName(Application.ExeName);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=caNone;
  LYTray1.HideMainForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  if (MessageDlg('�˳��󽫲��ٽ����豸����,ȷ���˳���', mtWarning, [mbYes, mbNo], 0) <> mrYes) then exit;
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

  ifSocketClient:=ini.readBool(IniSection,'Socket�ͻ���',false);//BC-10:�ͻ���
  ServerIP:=trim(ini.ReadString(IniSection,'������IP',''));
  ServerPort:=ini.ReadInteger(IniSection,'�������˿�',8080);//DH36��Ĭ�϶˿�Ϊ5600
  NoDtlStr:=ini.ReadInteger(IniSection,'������ʶλ',3);//BS300:4;DH36��BC-10:3

  autorun:=ini.readBool(IniSection,'�����Զ�����',false);
  ifRecLog:=ini.readBool(IniSection,'������־',false);
  EquipUnid:=ini.ReadInteger(IniSection,'�豸Ψһ���',-1);
  ifKLite8:=ini.readBool(IniSection,'KLite8��Ӧ',false);
  Line_Patient_ID:=ini.ReadString(IniSection,'������������','');
  No_Patient_ID:=ini.ReadInteger(IniSection,'������λ',3);
  CM_Category_Message:=ini.ReadString(IniSection,'��Ӧ��Ϣ����','');

  FS205_Chinese:=ini.readBool(IniSection,'�����������',false);
  BS300_Rerun:=ini.readBool(IniSection,'����BS300����',false);
  Discard_Qualitative:=ini.readBool(IniSection,'�������ŷָ��Ķ��Խ��',false);

  GroupName:=trim(ini.ReadString(IniSection,'������',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'������ĸ','')));//�������Ǵ�д������һʧ��
  SpecType:=ini.ReadString(IniSection,'Ĭ����������','');
  SpecStatus:=ini.ReadString(IniSection,'Ĭ������״̬','');
  CombinID:=ini.ReadString(IniSection,'�����Ŀ����','');

  LisFormCaption:=ini.ReadString(IniSection,'����ϵͳ�������','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9997');

  HisConnStr:=ini.ReadString(IniSection,'����HIS���ݿ�','');
  
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
  ss:='Socket�ͻ���'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '������IP'+#2+'Edit'+#2+#2+'1'+#2+'��λ��ͨ�Žӿڳ���Ϊ��������ʱ������д'+#2+#3+
      '�������˿�'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ʶλ'+#2+'Edit'+#2+#2+'1'+#2+'OBX���ô��߷ָ�,��0��ʼ,�ڼ�λ'+#2+#3+
      '������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ĸ'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '����ϵͳ�������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ����������'+#2+'Combobox'+#2+'<������д>'+#13+'OBR��15λ'+#2+'1'+#2+#2+#3+
      'Ĭ������״̬'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Ŀ����'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Զ�����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '������־'+#2+'CheckListBox'+#2+#2+'0'+#2+'ע:ǿ�ҽ�������������ʱ�ر�'+#2+#3+
      '�豸Ψһ���'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'KLite8��Ӧ'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '��Ӧ��Ϣ����'+#2+'Combobox'+#2+'ACK^R01'+#13+'ACK'+#2+'1'+#2+#2+#3+
      '������������'+#2+'Combobox'+#2+'PID'+#13+'OBR'+#2+'1'+#2+#2+#3+
      '������λ'+#2+'Edit'+#2+#2+'1'+#2+'PID��OBR���ô��߷ָ�,��0��ʼ,�ڼ�λ'+#2+#3+
      '�����������'+#2+'CheckListBox'+#2+#2+'1'+#2+'�ж�����:���ļ������ַ�(���)�Ƿ���ʾ����'+#2+#3+
      '����BS300����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '�������ŷָ��Ķ��Խ��'+#2+'CheckListBox'+#2+#2+'1'+#2+'iFlash-3000�����2.05,�޷�Ӧ�ԡ�'+#2+#3+
      '����HIS���ݿ�'+#2+'UniConn'+#2+#2+'1'+#2+'Oracle Server��ʽ:IP:Port:SID'+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2;

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
  //report_date:String;//����ʱ��
  //deptname:String;//�������
  //check_doctor:String;//����ҽ��
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
  r_check_doctor:String;//����ҽ��
begin
  if not if_test then Str:=Socket.ReceiveText;
  if FS205_Chinese then Str:=UTF8Decode(Str);//������ɲ�FS-205��������������
  
  if length(memo1.Lines.Text)>=1000000 then memo1.Lines.Clear;//memo��win98ֻ�ܽ���64K���ַ�,��win2000������
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

        //��HIS���ݿ�
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
            memo1.Lines.Add('����HIS���ݿ�ʧ��:'+E.Message);
          end;
        end;

        if not UniConnection1.Connected then begin UniConnection1.Free;continue;end;

        UniQuery1:=TUniQuery.Create(nil);
        UniQuery1.Connection:=UniConnection1;
        UniQuery1.SQL.Text:='select wm_concat(distinct ORDER_ID) as ORDER_ID from LIS_REQUEST where BARCODE='''+copy(Query_Target,Pos('^',Query_Target)+1,MaxInt)+''' ';//�ö�������HIS�����Ŀ����
        Try
          UniQuery1.Open;
        except
          on E:Exception do
          begin
            memo1.Lines.Add('��ѯHIS�����Ŀʧ��:'+E.Message);
          end;
        end;
        
        if not UniQuery1.Active then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        if UniQuery1.RecordCount<=0 then begin UniQuery1.Free;UniConnection1.Free;continue;end;

        aJson:=SO(GetLisCombItem(PChar(ConnectString),PChar(UniQuery1.fieldbyname('ORDER_ID').AsString),PChar(EquipChar),'PEIS'));

        if not aJson.AsObject.Exists('��Ŀ��Ϣ') then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        aSuperArray:=aJson['��Ŀ��Ϣ'].AsArray;
  
        ItemList:=TStringList.Create;
        for k:=0 to aSuperArray.Length-1 do
        begin
          ItemList.Add(aSuperArray[k]['�����Ŀ��ע'].AsString);//���������Ŀ�ġ�˵�����ֶΣ�1-ȫ����0-������2-�ɻ�ѧ
        end;

        //�����Ŀ,1-ȫ����0-������2-�ɻ�ѧ
        if (ItemList.IndexOf('1')>=0) or ((ItemList.IndexOf('0')>=0) and (ItemList.IndexOf('2')>=0)) then s1:='1'
          else if ItemList.IndexOf('0')>=0 then s1:='0'
            else if ItemList.IndexOf('2')>=0 then s1:='2'
              else begin ItemList.Free;UniQuery1.Free;UniConnection1.Free;continue;end;

        ItemList.Free;

        UniQuery1.Close;
        UniQuery1.SQL.Clear;
        UniQuery1.SQL.Text:='select top 1 * from LIS_REQUEST where BARCODE='''+copy(Query_Target,Pos('^',Query_Target)+1,MaxInt)+''' ';//�ö�������HIS�����Ŀ����
        Try
          UniQuery1.Open;
        except
          on E:Exception do
          begin
            memo1.Lines.Add('��ѯHIS�ܼ�����Ϣʧ��:'+E.Message);
          end;
        end;

        if not UniQuery1.Active then begin UniQuery1.Free;UniConnection1.Free;continue;end;
        if UniQuery1.RecordCount<=0 then begin UniQuery1.Free;UniConnection1.Free;continue;end;

        patientname:=UniQuery1.fieldbyname('NAME').AsString;//��������
        //�����Ա�
        if '1'=UniQuery1.fieldbyname('SEX').AsString then sex:='��'
          else if '2'=UniQuery1.fieldbyname('SEX').AsString then sex:='Ů'
            else if '3'=UniQuery1.fieldbyname('SEX').AsString then sex:='����';
        if UniQuery1.fieldbyname('AGEUNIT').AsString='Y' THEN AGEUNIT:='��'
          else if UniQuery1.fieldbyname('AGEUNIT').AsString='N' THEN AGEUNIT:='��';
        age:=UniQuery1.fieldbyname('AGE').AsString+AGEUNIT;//��������
        report_date:=UniQuery1.fieldbyname('WRITE_TIME').AsString;//����ʱ��
        deptname:=UniQuery1.fieldbyname('REQDEPT').AsString;//�������
        check_doctor:=UniQuery1.fieldbyname('WRITE_NAME').AsString;//����ҽ��
        His_Unid:=UniQuery1.fieldbyname('REG_ID').AsString;//����

        UniQuery1.Free;
        UniConnection1.Free;

        ORF:=#$0B+
             'MSH|^~\&|||||||ORF|1|P|2.3'+#$0D+
             'MSA|AA|'+Message_Control_ID+#$0D+
             'QRD||R|I||||20^LI|'+Query_Target+'|DEM|ALL'+#$0D+
             'PID|||'+Query_Target+'||'+s1+'|'+patientname+'|'+His_Unid+'|'+age+'|'+sex+#$0D+//PID-3����������ڱ�ʶ������ݵ�Ψһ��ʶ�š������������ź����롣�ǲ��Ƿ��ؽ���������ţ�Ϊ�ջ���ô����//PID-5����ģʽ��1��or��0��or��2��(ȫ�����������ɻ�ѧ)
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
        r_check_doctor:='';//����ҽ��

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
            if(SpecType='OBR��15λ')and(ls3.Count>15) then SpecType:=ls3[15];
            if ls3.Count>16 then r_check_doctor:=ls3[16];//����ҽ��
            ls3.Free;
          end;

          DtlStr:='';
          sValue:='';
          sHistogramFile:='';
          if uppercase(copy(trim(ls[i]),1,4))='OBX|' then
          begin
            ls2:=StrToList(ls[i],'|');

            if ls2.Count>NoDtlStr then DtlStr:=ls2[NoDtlStr];

            if(ls2.Count>5)and(ls2[2]<>'ED')then//ls2[2]='ED'��ʾͼƬ���
            begin
              sValue:=ls2[5];
              sValue:=StringReplace(sValue,'��','',[rfReplaceAll, rfIgnoreCase]);//�ɲ�FS-205
              sValue:=StringReplace(sValue,'��','',[rfReplaceAll, rfIgnoreCase]);//�ɲ�FS-205
              sValue:=StringReplace(sValue,'mg/L','',[rfReplaceAll, rfIgnoreCase]);//EU-5300
              sValue:=StringReplace(sValue,'mmol/L','',[rfReplaceAll, rfIgnoreCase]);//EU-5300

              //FUS2000
              ls7:=StrToList(sValue,'^');
              if ls7.Count>2 then sValue:=trim(ls7[1]+' '+ls7[2]);
              ls7.Free;

              //iFlash-3000.����,���Ϊ��2.05,�޷�Ӧ�ԡ�,ֻȡ����ǰ��2.05
              if Discard_Qualitative and(pos(',',sValue)>0) then sValue:=copy(sValue,1,pos(',',sValue)-1);

              //����EU-5300 begin
              if(pos(':\E\',sValue)>0)and(rightstr(sValue,4)='.JPG')then//��ʾ�����ͼƬ·��
              begin
                sHistogramFile:=StringReplace(sValue,'\E\','\',[rfReplaceAll]);
                sValue:='';
              end;
              //����EU-5300 end
            end;

            //ͼƬ���� strat
            if(ls2[2]='ED')and(ls2.Count>5)and(trim(ls2[5])<>'') then//ls2[2]='ED'��ʾͼƬ����,ls2[5]��ʾͼƬ����
            begin
              //DH36:ls2[5]Ϊ^Image^PNG^Base64^iVBORw0KGgoAAAANSUhEUgAAAJw.........
              //BC10:ls2[5]Ϊ^Image^BMP^Base64^Qk0GcAAAAAAAAL.........
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

              if pos('^',ls2[5])<=0 then//FUS-2000��GMD-S600��ls2[5]����ͼƬ����,û��^
              begin
                sHistogramFile:=DtlStr+ifThen(leftstr(ls2[5],4)='/9j/','.jpg','.bmp');//�μ��ĵ���MUSϵ��ȫ�Զ���Һ����ϵͳ�ӿڹ淶20220210.pdf��

                try
                  //FUS2000��ls2[5]ʵ���Ͽ��ܰ�������ͼƬ,��424D�ָ�,424D��sHistogramTemp��
                  //��LIS��֧�ֵ�����Ŀ����ͼƬ�ı�������ʾ,�պ�����Ĵ���ʽҲֻ��ʶ��һ��ͼƬ,�����ò�ִ�����
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
            //ͼƬ���� stop

            ls2.Free;
          end;
          ReceiveItemInfo[i]:=VarArrayof([DtlStr,sValue,'',sHistogramFile]);

          //�����������Start
          if BS300_Rerun then
          begin
            for  j:=0  to i-1 do
            begin
              if (DtlStr<>'')and(ReceiveItemInfo[j][0]=DtlStr) then ReceiveItemInfo[j]:=VarArrayof(['','','','']);
            end;
          end;
          //�����������End
        end;
        
        ls.Free;

        //������begin
        ls8:=StrToList(SpecNo,'|');
        if ls8.Count>No_Patient_ID then SpecNo:=ls8[No_Patient_ID];
        ls8.Free;
        SpecNo:=trim(StringReplace(SpecNo,'^R','',[rfReplaceAll, rfIgnoreCase]));//KLite8
        if SpecNo='' then SpecNo:=formatdatetime('nnss',now);
        SpecNo:=rightstr('0000'+SpecNo,4);
        //������end

        if bRegister then
        begin
          FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
          FInts.fData2Lis(ReceiveItemInfo,(SpecNo),CheckDate,
            (GroupName),(SpecType),(SpecStatus),(EquipChar),
            (CombinID),r_patientname+'{!@#}'+r_sex+'{!@#}{!@#}'+r_age+'{!@#}{!@#}{!@#}'+r_check_doctor,(LisFormCaption),(ConnectString),
            (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
            ifRecLog,true,'����',
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
