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
  HisConnStr:String;

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

  FS205_Chinese:=ini.readBool(IniSection,'�����������',false);
  BS300_Rerun:=ini.readBool(IniSection,'����BS300����',false);

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
      '������������'+#2+'Combobox'+#2+'PID'+#13+'OBR'+#2+'1'+#2+#2+#3+
      '������λ'+#2+'Edit'+#2+#2+'1'+#2+'PID��OBR���ô��߷ָ�,��0��ʼ,�ڼ�λ'+#2+#3+
      '�����������'+#2+'CheckListBox'+#2+#2+'1'+#2+'�ж�����:���ļ������ַ�(���)�Ƿ���ʾ����'+#2+#3+
      '����BS300����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
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
  i,j:integer;
  Str:string;
  SBPos,EBPos:integer;
  ls,ls2,ls3,ls4,ls5,ls7,ls8:tstrings;
  DtlStr:string;
  CheckDate:string;
  sHistogramTemp:string;
  sHistogramFile:string;
  strList:TStrings;
  Message_Control_ID:string;
  Query_Target:String;
  ORF:String;

  lls:TStrings;
  UniConnection1:TUniConnection;
  UniQuery1:TUniQuery;

  patientname:String;
  sex:String;
  age:String;
  report_date:String;//����ʱ��
  deptname:String;//�������
  check_doctor:String;//����ҽ��
begin
  if not if_test then Str:=Socket.ReceiveText;
  if FS205_Chinese then Str:=UTF8Decode(Str);//������ɲ�FS-205��������������
  
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

    ls:=TStringList.Create;
    ExtractStrings([#$D],[],Pchar(rfm2),ls);

    if (pos(#$0D'QRD|',rfm2)>0)and(pos(#$0D'QRF|',rfm2)>0) then//��������Ϣ
    begin
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

      //��HIS���ݿ�
      if trim(HisConnStr)<>'' then
      begin
        lls:=TStringList.Create;
        lls.Delimiter:=';';
        lls.DelimitedText:=HisConnStr;
      
        UniConnection1:=TUniConnection.Create(nil);
        UniConnection1.ProviderName:=lls.Values['ProviderName'];
        UniConnection1.SpecificOptions.Values['Direct']:='True';
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

        if UniConnection1.Connected then
        begin
          UniQuery1:=TUniQuery.Create(nil);
          UniQuery1.Connection:=UniConnection1;
          UniQuery1.SQL.Text:='select * from LIS_REQUEST';
          Try
            UniQuery1.Open;
          except
            on E:Exception do
            begin
              memo1.Lines.Add('�򿪱�ʧ��:'+E.Message);
            end;
          end;
          if UniQuery1.Active then
          begin
            //�����Ŀ
            patientname:=UniQuery1.fieldbyname('NAME').AsString;//��������
            //�����Ա�
            if '1'=UniQuery1.fieldbyname('SEX').AsString then sex:='��'
              else if '2'=UniQuery1.fieldbyname('SEX').AsString then sex:='Ů'; 
            age:=UniQuery1.fieldbyname('AGE').AsString+UniQuery1.fieldbyname('AGEUNIT').AsString;//��������
            report_date:=UniQuery1.fieldbyname('NAME').AsString;//����ʱ��
            deptname:=UniQuery1.fieldbyname('NAME').AsString;//�������
            check_doctor:=UniQuery1.fieldbyname('NAME').AsString;//����ҽ��
          end;
          UniQuery1.Free;
        end;

        UniConnection1.Close;
        UniConnection1.Free;
      end;

      ORF:=#$0B+
           'MSH|^~\&|||||||ORF|1|P|2.3'+#$0D+
           'MSA|AA|'+Message_Control_ID+#$0D+
           'QRD||R|I||||20^LI|'+Query_Target+'|DEM|ALL'+#$0D+
           'PID|||'+Query_Target+'||1||||'+#$0D+//PID-3����������ڱ�ʶ������ݵ�Ψһ��ʶ�š������������ź����롣�ǲ��Ƿ��ؽ���������ţ�Ϊ�ջ���ô����//PID-5����ģʽ��1��or��0��or��2��(ȫ�����������ɻ�ѧ)
           'PV1|||'+#$0D+
           'OBR||||||||||||||||'+#$0D+
           #$1C#$0D;
      Socket.SendText(ORF);
    end else//���ͼ����
    begin
      SpecNo:='';
          
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
        end;
      
        if uppercase(copy(trim(ls[i]),1,4))='OBR|' then
        begin
          if Line_Patient_ID='OBR' then SpecNo:=ls[i];

          ls3:=StrToList(ls[i],'|');
          if ls3.Count>7 then CheckDate:=copy(ls3[7],1,4)+'-'+copy(ls3[7],5,2)+'-'+copy(ls3[7],7,2)+' '+copy(ls3[7],9,2)+ifThen(copy(ls3[7],9,2)<>'',':')+copy(ls3[7],11,2);
          if(SpecType='OBR��15λ')and(ls3.Count>15) then SpecType:=ls3[15];
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

            //FUS2000
            ls7:=StrToList(sValue,'^');
            if ls7.Count>2 then sValue:=trim(ls7[1]+' '+ls7[2]);
            ls7.Free;
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
          (CombinID),'',(LisFormCaption),(ConnectString),
          (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
          ifRecLog,true,'����',
          '',
          EquipUnid,
          '','','','',
          -1,-1,-1,-1,
          -1,-1,-1,-1,
          false,false,false,false);
        if not VarIsEmpty(FInts) then FInts:= unAssigned;
      end;

      EBPos:=pos(#$1C#$0D,rfm);
    
      if ifKLite8 then
      begin
        //===================================ACK^R01===GMD-S600�������ΪACK.KLite8ʹ��ACK^R01ȷ��û����,�����ACK�ܷ�������KLite8
        Socket.SendText(#$0B+'MSH|^~$&|||||||ACK|1|P|2.4||||0||ASCII|||'+#$0D+'MSA|AA|'+Message_Control_ID+'|message accepted|||0|'+#$0D#$1C#$0D);
      end;
    end;
    ls.Free;
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
