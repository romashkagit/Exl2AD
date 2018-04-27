unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ActiveX,   // ������������ ��� COM Moniker stuff...
  ActiveDs_TLB,   // ��������� ���������� �����
  ComObj,// ������������ ��� OleCheck � ������ ������� COM
  StdCtrls, Vcl.ExtCtrls, Generics.Collections,ADODB, Data.DB, Vcl.ComCtrls,IniFiles;

type

  TUserInfo= record
    FIO, pswd, title, phone,room, company,department,manager,domens: string;
    status :Integer;
    class function Create(const FIO, pswd, title, phone, room, company,department, manager,domens: string; status:Integer): TUserInfo; static;

  end;

  TExl2ADfm = class(TForm)
    pnlBtn: TPanel;
    btnLoadExl: TButton;
    qCn: TADOQuery;
    adoCn: TADOConnection;
    reLog: TRichEdit;
    function get_wl(str: String; delimiter:char):TStringList;
   // procedure CreateUser(UserIsn :Double;ADName,Phone,Room,DeptName,DutyName,ClassName,Fullname,vStatus: string);
    procedure CreateUser(pair : TPair<string, TUserInfo>);//ADName,Fullname,Pswd,DutyName,Status: string);
    procedure UpdateUser(pair : TPair<string, TUserInfo>);//ADName,Fullname,Pswd,DutyName,Status,Phone,Room,Department,EmployeeType: string);
    function GetTranslitWord(s:String):string;
    function GetADName(Fullname:string;needadd:Integer=0):string;
    function GetObject(const Name: string): IDispatch;
    function GetAD_UserName(UserName: string; DomainName: string): string;
    procedure FormCreate(Sender: TObject);
    procedure LoadExl;
    procedure SaveLog(Msg :String;Color:TColor=clLime);
    procedure  ColorLine(RE : TRichEdit; Line : Integer; LineColor : TColor);
    procedure btnLoadExlClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Exl2ADfm: TExl2ADfm;
  UserDirectory : TDictionary<string, TUserInfo>;
  Fdomen, Fd1, Fd2:String;
implementation

{$R *.dfm}
class function TUserInfo.Create(const FIO, pswd, title, phone, room, company,department, manager,domens: string; status:Integer): TUserInfo;
begin
  Result.FIO := FIO;
  Result.pswd := pswd;
  Result.title := title;
  Result.phone := phone;
  Result.room := room;
  Result.company := company;
  Result.department := department;
  Result.manager := manager;
  Result.domens := domens;
  Result.Status := Status;
end;


procedure TExl2ADfm.SaveLog(Msg :String;Color:TColor=clLime);
var
  vSl:TStringList;
  vmsg,vpath:String;
begin
  vmsg := FormatDateTime('DD.MM.YYYY hh:mm:ss', now)+ ' ' + Msg;
  reLog.Lines.Insert(0,vmsg);
  ColorLine(reLog,0,Color);
  vpath :=  ExtractFilePath(Application.ExeName)+'Exl2AD.log';
  vSl:=TStringList.Create;
  try
  if FileExists(vpath) then
  vSl.LoadFromFile(vpath);
  vSl.Insert(0,char(13)+char(10)+vmsg);
  finally
    vSl.SaveToFile(vpath);
    vSl.Destroy;
  end;
end;

procedure  TExl2ADfm.ColorLine(RE : TRichEdit; Line : Integer; LineColor : TColor);
begin
  with RE do
  begin
    SelStart := SendMessage(Handle, EM_LINEINDEX, Line, 0);
    SelLength := Length(Lines[Line]);
    SelAttributes.Color := LineColor;
   // SelAttributes.Style := [fsBold];
  end;
end;

function TExl2ADfm.GetTranslitWord(s:String):string;
var
 i: integer;
 t: string;
begin
 for i:=1 to Length(s) do
  begin
   case s[i] of
        'a': t:=t+'a';
        '�': t:=t+'b';
        '�': t:=t+'v';
        '�': t:=t+'g';
        '�': t:=t+'d';
        '�': t:=t+'e';
        '�': t:=t+'e';
        '�': t:=t+'zh';
        '�': t:=t+'z';
        '�': t:=t+'i';
        '�': t:=t+'j';
        '�': t:=t+'k';
        '�': t:=t+'l';
        '�': t:=t+'m';
        '�': t:=t+'n';
        '�': t:=t+'o';
        '�': t:=t+'p';
        '�': t:=t+'r';
        '�': t:=t+'s';
        '�': t:=t+'t';
        '�': t:=t+'u';
        '�': t:=t+'f';
        '�': t:=t+'kh';
        '�': t:=t+'ts';
        '�': t:=t+'ch';
        '�': t:=t+'sh';
        '�': t:=t+'shh';
        '�': t:=t+'''';
        '�': t:=t+'y';
        '�': t:=t+'''';
        '�': t:=t+'e';
        '�': t:=t+'yu';
        '�': t:=t+'ya';
        '�': T:=T+'a';
        '�': T:=T+'b';
        '�': T:=T+'v';
        '�': T:=T+'g';
        '�': T:=T+'d';
        '�': T:=T+'e';
        '�': T:=T+'e';
        '�': T:=T+'zh';
        '�': T:=T+'z';
        '�': T:=T+'i';
        '�': T:=T+'y';
        '�': T:=T+'k';
        '�': T:=T+'l';
        '�': T:=T+'m';
        '�': T:=T+'n';
        '�': T:=T+'o';
        '�': T:=T+'p';
        '�': T:=T+'r';
        '�': T:=T+'s';
        '�': T:=T+'t';
        '�': T:=T+'u';
        '�': T:=T+'f';
        '�': T:=T+'kh';
        '�': T:=T+'ts';
        '�': T:=T+'ch';
        '�': T:=T+'sh';
        '�': T:=T+'shh';
        '�': T:=T+'''';
        '�': T:=T+'y';
        '�': T:=T+'''';
        '�': T:=T+'e';
        '�': T:=T+'yu';
        '�': T:=T+'ya';
      else t:=t+s[i];
   end;
  end;
 Result:=t;
end;

function TExl2ADfm.GetADName(Fullname:string;needadd:Integer=0):string;
var
FIO:TStrings;
vAdname:string;
begin
 FIO :=TStringList.Create;
 FIO.Delimiter :=' '; //����������� - ������
 FIO.DelimitedText :=Fullname; //����������� ����� �.�.�.
 vAdname :=GetTranslitWord(FIO[1][1])+'.';

 if needadd=0 then
 vAdname :=vAdname +GetTranslitWord(FIO[2][1])+'.';

 vAdname :=vAdname +GetTranslitWord(FIO[0])+'@'+FDomen;
 FIO.Destroy;
 Result:= vAdname;
end;

procedure TExl2ADfm.btnLoadExlClick(Sender: TObject);
var
 pair: TPair<string, TUserInfo>;
 vlogin:String;
begin
    UserDirectory.Clear;
    SaveLog('�������� �� Excel � �������:',clYellow);
    LoadExl;
    SaveLog('��������/���������� ������� �������',clYellow);
    //�������� �� ������ �������������
    //  UserDirectory.Add('r.ivanov',TUserInfo.Create('������ ����� ��������','12345','','','','',512));
    for pair  in UserDirectory do
    begin
    Application.ProcessMessages;
    vlogin:=GetAD_UserName(pair.Value.FIO, Fdomen);
    if  vlogin=pair.key then
     UpdateUser(pair)
    else
     CreateUser(pair);{UserDirectory.ExtractPair('r.ivanov'));}
    end;
    SaveLog('���������� ������ ������������� � AD ���������',clYellow);
    ShowMessage('���������� ������ ������������� � AD ���������!');
end;

procedure TExl2ADfm.LoadExl;
const
 arCollNames_List1 : Array [1..11,1..2] of string =
 (('A','������'),
  ('B','�����'),
  ('C','������'),
  ('D','���'),
  ('E', '���������'),
  ('F','�������'),
  ('G', '�������'),
  ('H','�����������'),
  ('I','�������������'),
  ('J','������������'),
  ('K','�������� �����')
  );

Function IColByName(lv_nameCol:String; pi_sgn_Doc:string):LongInt;
var
  i:LongInt;
begin
  Result:=0;
  if (uppercase(pi_sgn_Doc) = uppercase('Users'))  then
  For i:=Low(arCollNames_List1) to High(arCollNames_List1) do
  begin
    if AnsiUpperCase(arCollNames_List1[i,1]) = AnsiUpperCase(lv_nameCol) then
    begin
      Result:=i;
      Break;
    end;
  end;
end;


var Rows, Cols, i,j: integer;
    ExcelApp, WorkSheet: OLEVariant;
    d: TDateTime;
    openDialog : TOpenDialog;
    lv_NameWorkSheet, lv_sgn_Doc, lv_Extid: String;
    vFio,vlogin,vPswd,vtitle,vmanager, vdepartment, vphone,vroom,vcompany,vdomens,Stmp:String;
    lv_WorkSheetCount,vstatus,vempty : integer;
    lv_sgn_List1: boolean;
    EndLine, StartLine, iRow, iCol, HeadLine, iColIsLoad, iColErr, iColRemark, vLoadCount: LongInt;
begin
vempty:=0;
try
try
  UserDirectory.Clear;
  OpenDialog:=TOpenDialog.Create(Self);
  if OpenDialog.Execute then
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.DisplayAlerts := False;
  //��������� �����
  ExcelApp.Workbooks.Open({*ExtractFilePath(Application.ExeName)+'users.xlsm');//*}openDialog.FileName);
  lv_WorkSheetCount :=  ExcelApp.Worksheets.count;
  for i := 1 to lv_WorkSheetCount do
  begin
  ExcelApp.Worksheets[i].Activate;
  lv_NameWorkSheet :=  ExcelApp.ActiveSheet.Name;
  lv_sgn_List1  := uppercase(lv_NameWorkSheet) = uppercase('Users');
  if lv_sgn_List1 then
   begin
     HeadLine := 1;
     StartLine := 2;
     EndLine :=  ExcelApp.ActiveSheet.UsedRange.Rows.Count ;
     iColIsLoad :=High(arCollNames_List1)+4 ;
   end;
   if  lv_sgn_List1  then
    begin
    iColErr:=iColIsLoad+1;
    iColRemark := iColErr+1;

    end;

  // ��������� ������� ISLOAD (������� ��������)
    {if ExcelApp.ActiveSheet.Cells[HeadLine,iColIsLoad].Value<>'ISLOAD' then
    ExcelApp.ActiveSheet.Cells[HeadLine,iColIsLoad]:='ISLOAD';
   // ExcelApp.ActiveSheet.Cells[HeadLine,iColIsLoad]:=7;

    // ��������� ������� ERR (������)
    if ExcelApp.ActiveSheet.Cells[HeadLine,iColErr].Value<>'ERR' then
    ExcelApp.ActiveSheet.Cells[HeadLine,iColErr]:='ERR';
  //  ExcelApp.ActiveSheet.Cells[HeadLine,iColErr].ColumnWidth:=40;

    // ��������� ������ Remark (��� ����������)
    if ExcelApp.ActiveSheet.Cells[HeadLine,iColRemark].Value<>'Remark' then
    ExcelApp.ActiveSheet.Cells[HeadLine,iColRemark]:='Remark';
   // ExcelApp.ActiveSheet.Cells[HeadLine,iColRemark].ColumnWidth:  =60;  }

  For iRow:=StartLine to EndLine do
    begin
      sTmp:=ExcelApp.ActiveSheet.Cells[iRow,iColIsLoad].Value;
      if sTmp<>'' then Continue;
      // ���� ��������� �������� ��������������� �� ������, �� �� ���������
      if lv_sgn_List1  then
      begin
        //status
        sTmp:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('A',lv_NameWorkSheet)];
        if sTmp = 'R' then
        vstatus:=512
        else
        vstatus:=514;
        // LOGIN
        sTmp:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('B',lv_NameWorkSheet)];
        if sTmp <> '' then
        vlogin :=sTmp
        else
        begin
        Inc(vempty);
        if vempty>5 then
        Exit
        else
        Continue;
        end;
        // pswd
        vPswd:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('C',lv_NameWorkSheet)];
         // FIO
        vFIO:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('D',lv_NameWorkSheet)];
        // title
        vtitle:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('E',lv_NameWorkSheet)];
        //phone
        vphone:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('F',lv_NameWorkSheet)];
        //room
        vroom:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('G',lv_NameWorkSheet)];
        //company
        vcompany:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('H',lv_NameWorkSheet)];
        //department
        vdepartment:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('I',lv_NameWorkSheet)];
        //manager
        vmanager:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('J',lv_NameWorkSheet)];
        //domens
        vdomens:=ExcelApp.ActiveSheet.Cells[iRow,IColByName('K',lv_NameWorkSheet)];
        UserDirectory.Add(vlogin,TUserInfo.Create(vFIO,vPswd,vtitle,vphone,vroom,vcompany,vdepartment,vmanager,vdomens,vstatus));
        SaveLog('�������� �� xls � ������ ������������ '+vlogin);
      end;
  end;
  end;
  Except
  on E : Exception do
   SaveLog('������: '+ E.Message, clRed);
  end;
 finally
  SaveLog('�������� �� Excel � ������� ���������',clYellow);
  openDialog.Destroy;
  ExcelApp.ActiveWorkbook.Close;
 end;

end;

function TExl2ADfm.get_wl(str : String; delimiter:char):TStringList;
var sl: TStringList;
begin
  sl := TStringList.Create;
  sl.Delimiter := delimiter; //����������� - ������
  sl.DelimitedText :=str; //����������� ����� �.�.�.
  Result:=sl;
end;


procedure TExl2ADfm.CreateUser(pair : TPair<string, TUserInfo>);//ADName,Fullname,Pswd,DutyName,Status: string);
var
 Usr: IADsUser;
 Comp: IADsContainer;
 Value:TUserInfo;
 FIO, MailList: TStringList;
 i: integer;
begin
 try
 try
  Value:=pair.Value;
  Comp := GetObject('LDAP://CN=Users,DC='+Fd1+',DC='+Fd2) as  IADsContainer;
  Usr := Comp.Create('user','CN='+Value.FIO) as IADsUser;
  FIO :=TStringList.Create;
  FIO:=get_wl(Value.FIO,' ');
  for i:=0 to FIO.Count-1  do
  begin
      case i of
       0:  if FIO[0] <> '' THEN Usr.Put('sn',FIO[0] );     // �������
       1:  IF FIO[1] <> '' THEN  Usr.Put('givenName',FIO[1] );     // ��� ������������
       2:  IF FIO[1][1]+FIO[2][1]<> '' THEN Usr.Put('initials',FIO[1][1]+'.' +FIO[2][1]+'.');
      end;
  end;
  IF Value.FIO  <> '' THEN Usr.Put('displayName',Value.FIO );     // ��������� ���
  IF Value.title <> '' THEN  Usr.put('title',Value.title);
  Usr.put('userAccountControl',Value.status);  //������
  IF Value.phone <> '' THEN Usr.Put('telephoneNumber',Value.phone);   //�������
  IF Value.room <> '' THEN Usr.Put('physicalDeliveryOfficeName',Value.room);//�������
  IF Value.department <> '' THEN Usr.Put('department', Value.department);
  IF Value.company <> '' THEN Usr.Put('company', Value.company);
  IF Value.manager <> '' THEN Usr.Put('manager',Value.manager);
  if Value.domens<> '' THEN
  begin
  MailList:=TStringList.Create;
  MailList:=get_wl(StringReplace(Value.domens, '''', ' ',[rfReplaceAll, rfIgnoreCase]),',');
  for i := 0 to MailList.Count-1 do
  begin
  MailList[i]:=  LowerCase(pair.Key)+MailList[i];
  if i>0 then   MailList[i]:=', '+ MailList[i];
  end;
  Usr.Put('mail',MailList.Text);
  end;

  Usr.put('sAMAccountName',pair.Key);
 //�������� ���� �� ����� ������ � ActiveDirectory
 //if (GetAD_UserName(GetADName(Fullname,0),'userPrincipalName')='') then
  Usr.put('userPrincipalName',pair.Key);//GetADName(Fullname,0));
 // else   //���� ����, �� ��������� ��������
 // Usr.put('userPrincipalName',GetADName(Fullname,1));
  Usr.SetInfo;
  IF Value.pswd <> '' THEN Usr.SetPassword(Value.pswd);
  SaveLog('������������ '+Value.FIO+' ������ � Active Directory');
 except
  on E: EOleException do begin
  SaveLog('������ ��� �������� ������������ '+Value.FIO+': '+E.Message, clRed);
  end;
 end;
 finally
    FIO.Free;
    MailList.Free;
 end;

end;

procedure TExl2ADfm.UpdateUser(pair : TPAir<string, TUserInfo>);//ADName,Fullname,Pswd,DutyName,Status,Phone,Room,Department,EmployeeType: string);
var
 Usr: IADsUser;
 Comp: IADsContainer;
 Value:TUserInfo;
 FIO, MailList: TStringList;
  i: integer;
begin
 try
 try
  Value:=pair.Value;
  Comp := GetObject('LDAP://CN=Users,DC='+Fd1+',DC='+Fd2) as  IADsContainer;
  Usr := Comp.GetObject('user','CN='+Value.FIO)  as IADsUser;
  FIO:=TStringList.Create;
  FIO:=get_wl(Value.FIO,' ');
  for i:=0 to FIO.Count-1  do
  begin
      case i of
       0:  if FIO[0] <> '' THEN Usr.Put('sn',FIO[0] );     // �������
       1:  IF FIO[1] <> '' THEN  Usr.Put('givenName',FIO[1] );     // ��� ������������
       2:  IF FIO[1][1]+FIO[2][1]<> '' THEN Usr.Put('initials',FIO[1][1]+'.' +FIO[2][1]+'.');
      end;
  end;
  IF Value.FIO  <> '' THEN Usr.Put('displayName',Value.FIO );     // ��������� ���
  IF Value.title <> '' THEN  Usr.put('title',Value.title);
  Usr.put('userAccountControl',Value.status);  //������
  IF Value.phone <> '' THEN Usr.Put('telephoneNumber',Value.phone);   //�������
  IF Value.room <> '' THEN Usr.Put('physicalDeliveryOfficeName',Value.room);//�������
  IF Value.department <> '' THEN Usr.Put('department', Value.department);
  IF Value.company <> '' THEN Usr.Put('company', Value.company);
  IF Value.manager <> '' THEN Usr.Put('manager',Value.manager);
  if Value.domens<> '' THEN
  begin
  MailList:= TStringList.Create;
  MailList:=get_wl(StringReplace(Value.domens, '''', ' ',[rfReplaceAll, rfIgnoreCase]),',');
 // ShowMessage(MailList.Text);
  for i := 0 to MailList.Count-1 do
  begin
  MailList[i]:= LowerCASE(pair.Key)+MailList[i];
  if i>0 then   MailList[i]:=', '+ MailList[i];
  end;
 // ShowMessage(MailList.Text);
  Usr.Put('mail',MailList.Text);
  end;


  Usr.put('sAMAccountName',pair.Key);
    // Usr:=Comp.MoveHere('LDAP://CN='+vOldValue+',DC='+Fd1+',DC='+Fd2',vNewValue) as IADsUser ;
     Usr.SetInfo;
  SaveLog('������������ '+pair.Value.FIO+' ������� � Active Directory');
 except
  on E: EOleException do begin
  SaveLog('������ ��� ��������� ������������ '+Value.FIO+': ' + E.Message, clRed);
  end;
 end;
 finally
   FIO.Free;
   MailList.Free;
 end;

end;

//�������� ����������� ������������ � ������ � Acrive Directory
function TExl2ADfm.GetAD_UserName(UserName: string; DomainName: string): string;
begin
  Result := '';
  with qCn do
  begin
    Close;
    SQL.Clear;
    SQL.Add('select userPrincipalName, cn, department, title from ''LDAP://'+DomainName+''' where cn='''+UserName+'''');
  end;

  try
    if not qCn.Active then qCn.Open;
    Result:=qCn.FieldByName('userPrincipalName').AsString;
    if pos('@', Result)>0 then
    begin
      Result:=copy(Result, 1, pos('@', Result) -1);
    end;
  except
  end;
end;

procedure TExl2ADfm.FormCreate(Sender: TObject);
var Init1:TIniFile;
d:TStringList;
begin
      UserDirectory:=TDictionary<string, TUserInfo>.Create();
      //������
      FileSetAttr(ExtractFilePath(Application.ExeName)+'Exl2AD.ini', faHidden);
      Init1:= TIniFile.Create(ExtractFilePath(Application.ExeName)+'Exl2AD.ini');
      Fdomen:= init1.ReadString('Domen info','Domen', '');
      d:=TStringList.Create;
      d:=get_wl(Fdomen,'.');
      Fd1:=d[0];
      Fd2:=d[1];
end;


function TExl2ADfm.GetObject(const Name: string): IDispatch;

var

 Moniker: IMoniker;

 Eaten: integer;

 BindContext: IBindCtx;

 Dispatch: IDispatch;

begin

 OleCheck(CreateBindCtx(0, BindContext));

 OleCheck(MkParseDisplayName(BindContext,

  PWideChar(WideString(Name)),

  Eaten,

  Moniker));

 OleCheck(Moniker.BindToObject(BindContext, NIL, IDispatch,

  Dispatch));

 Result := Dispatch;

end;



end.
