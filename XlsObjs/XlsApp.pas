unit XlsApp;

interface

uses
  Classes, SysUtils, Excel97,
  XlsObjs;

type
  IXlsApplication = Excel97.ExcelApplication;
  TXlsAppObj = Excel97.TExcelApplication;

  TXlsApplication = class
  private
    FApp: TXlsAppObj;
    FXlsBooks: TXlsWorkbooks;
    FOwner: TComponent;

    procedure CreateExcel;
    procedure ReleaseExcel;
    function GetAsExcelApplication: IXlsApplication;
  protected

  public
    constructor Create( AOwner: TComponent );
    destructor Destroy; override;

    property Owner: TComponent read FOwner write FOwner;
    property IExcelApp: TXlsAppObj read FApp;
    property XlsBooks: TXlsWorkbooks read FXlsBooks;

    procedure ShowExcel;
    procedure HideExcel;
  end;

implementation

uses
  ActiveX, OleServer;

{ TXlsApplication }

constructor TXlsApplication.Create( AOwner: TComponent );
begin
  FOwner        := AOwner;
  FXlsBooks     := nil;
  FApp          := nil;

  CreateExcel;
end;

procedure TXlsApplication.CreateExcel;
begin
  FApp := Excel97.TExcelApplication.Create( FOwner );

  Assert( FApp <> nil );

//  FApp.ConnectKind := ckRunningOrNew;
//  FApp.DisplayAlerts[ DEF_LCID ] := False;
//  FApp.ConnectKind := ckNewInstance;

  FApp.Connect;

  if FApp.Application = nil then
    raise Exception.Create( 'Can not connect to excel out proc server.' );
  FXlsBooks := TXlsWorkbooks.Create( IExcelApp.Workbooks );
end;

destructor TXlsApplication.Destroy;
begin
  FreeAndNil( FXlsBooks );

  ReleaseExcel;
  inherited;
end;

function TXlsApplication.GetAsExcelApplication: IXlsApplication;
begin
  Assert( FApp <> nil );
  Result := FApp.Application;
end;

procedure TXlsApplication.HideExcel;
begin
  Assert( FApp <> nil );

  FApp.Visible[ DEF_LCID ] := False;
end;

procedure TXlsApplication.ReleaseExcel;
begin
  Assert( FApp <> nil );

  if FApp.Application <> nil then
  begin
    if FApp.Workbooks.Count = 0 then
    begin
      FApp.Quit;
    end
    else
    begin
      FApp.WindowState[ DEF_LCID ]       := TOLEEnum(xlMinimized);
      FApp.Visible[ DEF_LCID ]           := True;
    end;
  end;

  FreeAndNil( FApp );
end;

procedure TXlsApplication.ShowExcel;
begin
  Assert( FApp <> nil );

  FApp.Visible[ DEF_LCID ] := True;
  if FApp.WindowState[ DEF_LCID ] = TOLEEnum(xlMinimized) then
    FApp.WindowState[ DEF_LCID ] := TOLEEnum(xlNormal);
  FApp.ScreenUpdating[ DEF_LCID ] := True;
end;

end.
