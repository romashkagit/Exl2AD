unit uLogManager;


interface

uses
  Windows, SysUtils, SyncObjs,
  uSingletonTemplate;

type
  TLogMsgType = (lmtInfo, lmtWarning, lmtError);

  TLogMsgLevel = (lmlMax, lmlInfo, lmlDebug);

  TLogManagerEvent = procedure(const aLogMsgText: string; aLogMsgType: TLogMsgType);
  TLogManagerEventObj = procedure(const aLogMsgText: string; aLogMsgType: TLogMsgType) of object;

  TLogManager = class(TSingleton)
  private
    fcs   : TCriticalSection;
    fFileName     : string; // файл лога
    fMaxLogLevel  : TLogMsgLevel; // уровень, который опредеяет записывать сообщение или нет
    fLogEvent     : boolean; // вызывать свойства?
    fOnLogEvent   : TLogManagerEvent;
    fOnLogEventObj  : TLogManagerEventObj;
    function FormatLogTime(aDT: TDateTime): string;
    function FormatLogMsgType(aLogMsgType: TLogMsgType): string;
    function FormatLogMsgText(aText: string; aLogMsgType: TLogMsgType): string;
    procedure DoLogEvent(aLogMsgText: string; aLogMsgType: TLogMsgType; aWriteToEvent: boolean);
  protected
    constructor Create; override;
  public
    procedure WriteToLog(aText: string; aLogMsgType: TLogMsgType = lmtInfo; aLvl: TLogMsgLevel = Low(TLogMsgLevel);
      aWriteToEvent: boolean = True);
  public
    destructor Destroy; override;
    property FileName: string read fFileName write fFileName; // куда писать
    property MaxLogLevel: TLogMsgLevel read fMaxLogLevel write fMaxLogLevel; // максимальный доступный уровень
    property LogEvent: boolean read fLogEvent write fLogEvent; // вызывать события лога?
    property OnLogEvent: TLogManagerEvent read fOnLogEvent write fOnLogEvent; // событие лога
    property OnLogEventObj: TLogManagerEventObj read fOnLogEventObj write fOnLogEventObj; // событие лога для объектов
  end;

// funcs
  function LogMng: TLogManager;

implementation

function LogMng: TLogManager;
begin
  Result := TLogManager.GetInstance;
end;

{ TLogManager }

constructor TLogManager.Create;
begin
  inherited;
  fcs := TCriticalSection.Create;
  fFileName := '';
  fMaxLogLevel := High(TLogMsgLevel);
  fLogEvent := False;
end;

destructor TLogManager.Destroy;
begin
  fcs.Enter;
  try
  finally
    fcs.Leave;
  end;
  fcs.Free;
  inherited;
end;

function TLogManager.FormatLogTime(aDT: TDateTime): string;
begin
  DateTimeToString(Result, '[dd.mm.yyyy] [hh:mm:ss]', aDT);
end;

function TLogManager.FormatLogMsgType(aLogMsgType: TLogMsgType): string;
begin
  case aLogMsgType of
    lmtInfo: Result := '[Info]';
    lmtWarning: Result := '[Warning]';
    lmtError: Result := '[Error]';
  end;
end;

function TLogManager.FormatLogMsgText(aText: string; aLogMsgType: TLogMsgType): string;
begin
  Result := format('%s %s: %s', [FormatLogTime(Now), FormatLogMsgType(aLogMsgType), aText]);
end;

procedure TLogManager.DoLogEvent(aLogMsgText: string; aLogMsgType: TLogMsgType; aWriteToEvent: boolean);
begin
  if not fLogEvent or not aWriteToEvent then
    Exit;
  if Assigned(fOnLogEvent) then
    fOnLogEvent(aLogMsgText, aLogMsgType);
  if Assigned(fOnLogEventObj) then
    fOnLogEventObj(aLogMsgText, aLogMsgType);
end;

procedure TLogManager.WriteToLog(aText: string; aLogMsgType: TLogMsgType; aLvl: TLogMsgLevel; aWriteToEvent: boolean);
var
  fh  : THandle;
  lmsg  : string;
  card  : cardinal;
begin
  if fFileName = '' then
  begin
    lmsg := FormatLogMsgText(aText, aLogMsgType);
    lmsg := Format('%s: %s', [FormatLogMsgType(lmtError), 'Не определён файл логирования! LogMsg: ' + lmsg]);
    DoLogEvent(lmsg, lmtError, True);
    Exit;
  end;
  if aLvl > fMaxLogLevel then
    Exit;
  fh := CreateFile(PChar(fFileName), GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE,
    nil, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
  if fh = INVALID_HANDLE_VALUE then
  begin
    lmsg := FormatLogMsgText(aText, aLogMsgType);
    lmsg := Format('%s: %s', [FormatLogMsgType(lmtError), 'Ошибка лог-файла: INVALID_HANDLE_VALUE. LogMsg: ' + lmsg]);
    DoLogEvent(lmsg, lmtError, True);
    Exit;
  end;
  fcs.Enter;
  try
    // формируем сообщение
    lmsg := FormatLogMsgText(aText, aLogMsgType);
    DoLogEvent(lmsg, aLogMsgType, aWriteToEvent);
    lmsg := lmsg + #13#10;
    try
      SetFilePointer(fh, 0, nil, FILE_END);
      WriteFile(fh, lmsg[1], sizeof(lmsg[1]) * length(lmsg), card, nil);
    except
      on E: Exception do
      begin
        lmsg := Format('%s: %s', [FormatLogMsgType(lmtError), 'Ошибка записи в файл лога: ' +
          E.ClassName + ': ' + E.Message]);
        DoLogEvent(lmsg, lmtError, True);
      end;
      Else
      begin
        lmsg := Format('%s: %s', [FormatLogMsgType(lmtError), 'Неизвестная ошибка в TLogManager.WriteToLog!']);
        DoLogEvent(lmsg, lmtError, True);
      end;
    end;
  finally
    CloseHandle(fh);
    fcs.Leave;
  end;
end;

end.
