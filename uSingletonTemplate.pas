unit uSingletonTemplate;

interface

type
  TSingleton = class(TObject)
  private
    class procedure RegisterInstance(aInstance: TSingleton);
    procedure UnRegisterInstance;
    class function FindInstance: TSingleton;
  protected
    constructor Create; virtual;
  public
    class function NewInstance: TObject; override;
    procedure BeforeDestruction; override;
    constructor GetInstance; // Точка доступа к экземпляру
  end;

implementation

uses
  Contnrs;

var
  SingletonList : TObjectList;

{ TSingleton }

class procedure TSingleton.RegisterInstance(aInstance: TSingleton);
begin
   SingletonList.Add(aInstance);
end;

procedure TSingleton.UnRegisterInstance;
begin
   SingletonList.Extract(Self);
end;

class function TSingleton.FindInstance: TSingleton;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to SingletonList.Count - 1 do
    if SingletonList[i].ClassType = Self then
  begin
    Result := TSingleton(SingletonList[i]);
    Break;
  end;
end;

constructor TSingleton.Create;
begin
  inherited Create;
end;

class function TSingleton.NewInstance: TObject;
begin
  Result := FindInstance;
  if Result = nil then
  begin
    Result := inherited NewInstance;
    TSingleton(Result).Create;
    RegisterInstance(TSingleton(Result));
  end;
end;

procedure TSingleton.BeforeDestruction;
begin
   UnregisterInstance;
   inherited BeforeDestruction;
end;

constructor TSingleton.GetInstance;
begin
   inherited Create;
end;

initialization
   SingletonList := TObjectList.Create(True);

finalization
   SingletonList.Free;

end.
