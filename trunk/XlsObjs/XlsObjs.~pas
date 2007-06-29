unit XlsObjs;

interface

uses
  Classes, Contnrs, Excel97, Windows, Ranger;

const
  DEF_LCID      = 0;
  MY_COLORS     = 25;

type
  IXlsRange             = Range;
  IXlsWorkbook          = ExcelWorkbook;
  IXlsWorkbooks         = Workbooks;
  IXlsSheet             = ExcelWorksheet;
  IXlsSheets            = Sheets;

  TXlsWorksheet         = class;
  TXlsWorksheets        = class;

  TXlsRange             = class;
  TXlsRanges            = class;

  TXlsItemObj = class
  private
    FItemDisp: IDispatch;
    procedure SetItemDisp(const Value: IDispatch); virtual;
    function GetCompareName( ): string; virtual; abstract;
  public
    constructor Create( ); virtual;
    destructor Destroy; override;

    property ItemDisp: IDispatch read FItemDisp write SetItemDisp;
  end;

  CXlsItemObj = class of TXlsItemObj;

  TXlsObjContainer = class(TObjectList)
  private
    FContnrDisp: IDispatch;
  protected
    procedure DoDeleteItem( AItem: TObject ); virtual;
    procedure DoAddItem( AItem: TObject ); virtual;

    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
  public
    constructor Create( AContnr: IDispatch ); virtual;
    destructor Destroy(); override;

    property ContnrDisp: IDispatch read FContnrDisp write FContnrDisp;

    function AddItem: TXlsItemObj; virtual; abstract;
    function GetByName( AName: string ): TXlsItemObj;

    procedure RemoveItem( AName: string );

    class function CreateItem( AItemRef: CXlsItemObj; ADisp: IDispatch ): TXlsItemObj;
  end;

  TXlsWorkbook = class(TXlsItemObj)
  private
    FXlsWorksheets: TXlsWorksheets;
    FWasOpened: boolean;
    function GetAsWorkbook: IXlsWorkbook;
    procedure SetItemDisp(const Value: IDispatch); override;
    function GetCompareName( ): string; override;
  public
    constructor Create( ); override;
    destructor Destroy; override;

    property IWorkbook: IXlsWorkbook read GetAsWorkbook;
    property XlsWorksheets: TXlsWorksheets read FXlsWorksheets;
    property WasOpened: boolean read FWasOpened write FWasOpened;

    procedure Close( ASaveChanges: boolean );
  end;

  TXlsWorkbooks = class(TXlsObjContainer)
  private
    FOpenedCount: integer;
    function GetAsWorkbooks: IXlsWorkbooks;

    procedure GetBookSheets( ABook: TXlsWorkbook );
    procedure GetOpenedBooks;
    function GetByNameFromWorkbooks( AName: string ): TXlsWorkbook;
  protected
    procedure DoDeleteItem( AItem: TObject ); override;
    procedure DoAddItem( AItem: TObject ); override;
  public
    constructor Create( AContnr: IDispatch ); virtual;

    property IWorkbooks: IXlsWorkbooks read GetAsWorkbooks;

    //Create new item and add into container;
    function AddItem: TXlsItemObj; override;
    //Open excel file and add into container
    function OpenFile( AFileName: string ): TXlsItemObj;
    procedure CloseAll;

    function GetOpenedCount: integer;

    function GetAsIWorkbook( AName: string ): IXlsWorkbook;
    //Looking in container and then in the collection
    function FindBook( AName: string ): TXlsWorkbook;
  end;

  TXlsWorksheet = class(TXlsItemObj)
  private
    FRanges: TXlsRanges;
    function GetAsWorksheet: IXlsSheet;
    procedure SetItemDisp(const Value: IDispatch); override;
    function GetCompareName( ): string; override;
  public
    constructor Create(); override;
    destructor Destroy; override;

    property IWorksheet: IXlsSheet read GetAsWorksheet;
    property XlsRanges: TXlsRanges read FRanges;

    //GetSheetRange functions not create XlsRange object and get it directly
    //from ExcelWoorksheet object. It's not using XlsRanges object
    //A1:B30, A1:A1
    function GetSheetRange( AFrom, ATo: string ): IXlsRange; overload;
    //A1-R1C1
    function GetSheetRange( ARow, ACol: integer ): IXlsRange; overload;
    //ByName
    function GetSheetRange( ARngName: string ): IXlsRange; overload;

    procedure Delete;
  end;

  TXlsWorksheets = class(TXlsObjContainer)
  private
    function GetAsWorksheets: IXlsSheets;
  protected

  public
    property IWorksheets: IXlsSheets read GetAsWorksheets;
    //Create new item and add into container; Add new sheet at the end;
    function AddItem: TXlsItemObj; override;

    function GetAsIWoorksheet( AName: string ): IXlsSheet;
  end;

  TXlsRange = class(TXlsItemObj)
  private
    function GetAsRange: IXlsRange;
    //Return range name or not absolute adress
    function GetCompareName( ): string; override;
  public
    property IRange: IXlsRange read GetAsRange;
  end;

  TXlsRanges = class(TXlsObjContainer)
  private
    //Create and add into container if find.
    function GetByNameFromSheet( AName: string ): TXlsRange;
  protected
  public
    //Do nothing
    function AddItem: TXlsItemObj; override;

    //Create and add relative range object -->> A1:B30, A1:A1
    function Add( ARngName, AFrom, ATo: string ): TXlsRange; overload;
    //Create and add absolute range object -->> A1-R1C1
    function Add( ARngName: string; ARow, ACol: integer ): TXlsRange; overload;
    function Add( ARngName, AAdrress: string ): TXlsRange; overload;


    //Seek item in the container, if not found in the excel file.
    //Return nil if not exist
    function FindRange( AName: string ): TXlsRange;
    //Same as above just return as interface
    function FindAsIRange( AName: string ): IXlsRange;
  end;

  function AbsoluteColToRef( ACol: integer ): string;
  function AbsoluteToRef( ACol, ARow: integer ): string;

  {
  LOCALE_USER_DEFAULT
//GetUserDefaultLCID
//, Windows
}
implementation

uses
  SysUtils {$IFDEF VER140}, Variants {$ENDIF};
  
const
  DEF_ADD_SHEETS        = 1;



function AbsoluteColToRef( ACol: integer ): string;
const
  ALF_COUNT      = 26;
  FIRST_LETTER   = 'A';
var
  AInd: integer;
  Base: integer;
  BaseMod: integer;
  BaseDiv: integer;
begin
  Assert( ( ACol > 0 ) and ( ACol <= 256) );

  Result := '';
  AInd := Ord( FIRST_LETTER );
  AInd := AInd - 1;
  Base := ACol;
  while Base > 0 do
  begin
    BaseDiv := Base div ALF_COUNT;
    BaseMod := Base mod ALF_COUNT;
    if BaseDiv = 0 then
    begin
      Result    := Result + Chr( AInd + ( BaseMod ) );
      Base      := 0;
    end
    else
    begin
      Result    := Result + Chr( AInd + ( BaseDiv ) );
      Base      := BaseMod;
    end;
  end;
end;

function AbsoluteToRef( ACol, ARow: integer ): string;
var
  AbsCol: string;
begin
  Assert( ( ARow > 0 ) and ( ARow < 65536 ) );
  Result := '';

  Result := AbsoluteColToRef( ACol ) + IntToStr( ARow );
end;


//A1:B30, A1:A1
function GetRange( ASheet: IXlsSheet; AFrom, ATo: string ): IXlsRange; overload;
begin
  Result := ASheet.Range[ AFrom, ATo ];
end;

//A1-R1C1
function GetRange( ASheet: IXlsSheet; ARow, ACol: integer ): IXlsRange; overload;
begin
  Result := IDispatch(ASheet.Cells.Item[ ARow, ACol ]) as IXlsRange;
end;

//ByName
function GetRange( ASheet: IXlsSheet; ARngName: string ): IXlsRange; overload;
begin
  Result := nil;
  try
    Result := ASheet.Range[ ARngName, EmptyParam ];
  except
    //expected exception. Excel raise exception then there is no range with name.
  end;
end;

{ TXlsObjContainer }

constructor TXlsObjContainer.Create(AContnr: IDispatch);
begin
  inherited Create( True );
  ContnrDisp := AContnr;
end;

class function TXlsObjContainer.CreateItem(AItemRef: CXlsItemObj;
  ADisp: IDispatch): TXlsItemObj;
var
  Item: TXlsItemObj;
begin
  Assert( AItemRef <> nil );
  Assert( ADisp <> nil );

  Result := nil;

  Item := nil;
  Item := AItemRef.Create;
  try
    Item.ItemDisp := ADisp;
  except
    if Item <> nil then
      Item.Free;
    raise;
  end;

  Result := Item;
end;

destructor TXlsObjContainer.Destroy;
begin
  FContnrDisp := nil;
  inherited;
end;

procedure TXlsObjContainer.DoAddItem(AItem: TObject);
begin

end;

procedure TXlsObjContainer.DoDeleteItem(AItem: TObject);
begin
  TXlsItemObj(AItem).ItemDisp := nil;
  if Count = 0 then
    ContnrDisp := nil;
end;


function TXlsObjContainer.GetByName(AName: string): TXlsItemObj;
var
  i: integer;
  ItemName: string;
begin
  Result := nil;

  i := 0;
  while i < Count do
  begin
    ItemName := TXlsItemObj(Items[ i ]).GetCompareName();
    if ItemName = AName then
    begin
      Result := TXlsItemObj( Items[ i ] );
      i := Count;
    end;
    Inc( i );
  end;
end;

procedure TXlsObjContainer.Notify(Ptr: Pointer; Action: TListNotification);
begin
  Assert( OwnsObjects );

  case Action of
    lnDeleted   : DoDeleteItem( TObject(Ptr) );
    lnAdded     : DoAddItem( TObject(Ptr) );
  end;

  inherited Notify(Ptr, Action);
end;

procedure TXlsObjContainer.RemoveItem(AName: string);
var
  RemoveItm: TXlsItemObj;
begin
  RemoveItm := GetByName( AName );
  if RemoveItm <> nil then
    Remove( RemoveItm );
end;

{ TXlsWorkbook }

procedure TXlsWorkbook.Close(ASaveChanges: boolean);
begin
  IWorkbook.Close( ASaveChanges, EmptyParam, EmptyParam, DEF_LCID );
end;

constructor TXlsWorkbook.Create( );
begin
  inherited;

  WasOpened             := False;
  FXlsWorksheets        := nil;

  FXlsWorksheets := TXlsWorksheets.Create( nil );
end;

destructor TXlsWorkbook.Destroy;
begin
  FXlsWorksheets.Free;
  FXlsWorksheets := nil;
  inherited;
end;

function TXlsWorkbook.GetAsWorkbook: IXlsWorkbook;
begin
  Result := nil;
  if ItemDisp <> nil then
    Result := ItemDisp as IXlsWorkbook;
end;

function TXlsWorkbook.GetCompareName: string;
begin
  Result := IWorkbook.FullName[ DEF_LCID ];
end;

procedure TXlsWorkbook.SetItemDisp(const Value: IDispatch);
begin
  inherited SetItemDisp( Value );

  Assert( FXlsWorksheets <> nil );

  if IWorkbook <> nil then
  begin
    FXlsWorksheets.ContnrDisp := IWorkbook.Sheets
  end
  else
  begin
    FXlsWorksheets.Clear;
  end;

end;

{ TXlsWorkbooks }

function TXlsWorkbooks.AddItem: TXlsItemObj;
var
  Book: TXlsItemObj;
begin
  Result := nil;

  Book := CreateItem( TXlsWorkbook, IWorkbooks.Add( EmptyParam, DEF_LCID ) );

  Assert( Book <> nil );

  Add( Book );

  Result := Book;
end;

procedure TXlsWorkbooks.CloseAll;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
  begin
    TXlsWorkbook(Items[i]).Close( False );
  end;

  Clear;

  FOpenedCount := 0;
end;

constructor TXlsWorkbooks.Create(AContnr: IDispatch);
begin
  inherited Create( AContnr );
  FOpenedCount := 0;

  GetOpenedBooks;
end;

procedure TXlsWorkbooks.DoAddItem(AItem: TObject);
begin
  Assert( AItem <> nil );
  GetBookSheets( TXlsWorkbook(AItem) );
end;

procedure TXlsWorkbooks.DoDeleteItem(AItem: TObject);
begin
  inherited;
end;

function TXlsWorkbooks.FindBook(AName: string): TXlsWorkbook;
begin
  Result := nil;

  Result := TXlsWorkbook(GetByName( AName ));

  if Result = nil then
    Result := GetByNameFromWorkbooks( AName );
end;

function TXlsWorkbooks.GetAsIWorkbook(AName: string): IXlsWorkbook;
var
  Book: TXlsItemObj;
begin
  Result := nil;

  Book := nil;
  Book := GetByName( AName );

  if Book <> nil then
    Result := TXlsWorkbook(Book).IWorkbook;
end;

function TXlsWorkbooks.GetAsWorkbooks: IXlsWorkbooks;
begin
  Assert( ContnrDisp <> nil );

  Result := ContnrDisp as IXlsWorkbooks;
end;

procedure TXlsWorkbooks.GetBookSheets( ABook: TXlsWorkbook );
var
  i: integer;
  Sh: TXlsItemObj;
begin
  Assert( ABook <> nil );
  with ABook.IWorkbook do
  begin
    for i := 1 to Sheets.Count do
    begin

      Sh := CreateItem( TXlsWorksheet, Sheets.Item[ i ] );

      Assert( Sh <> nil );

      ABook.XlsWorksheets.Add( Sh );
    end;
  end;
end;

function TXlsWorkbooks.GetByNameFromWorkbooks(AName: string): TXlsWorkbook;
var
  i: integer;
  Book: TXlsItemObj;
begin
  Result := nil;

  for i := 1 to IWorkbooks.Count do
  begin
    if IWorkbooks.Item[ i ].Name = AName then
    begin
      Book := CreateItem( TXlsWorkbook, IWorkbooks.Item[ i ] );

      Assert( Book <> nil );
      Add( Book );

      Result := TXlsWorkbook(Book);
    end;
  end;
end;

procedure TXlsWorkbooks.GetOpenedBooks;
var
  i: integer;
  Book: TXlsItemObj;
begin
  for i := 1 to IWorkbooks.Count do
  begin
    Book := CreateItem( TXlsWorkbook, IWorkbooks.Item[ i ] );
    TXlsWorkbook(Book).WasOpened := True;
    Add( Book );
    Inc( FOpenedCount );
  end;
end;

function TXlsWorkbooks.GetOpenedCount: integer;
begin
  Result := FOpenedCount;
end;

function TXlsWorkbooks.OpenFile(AFileName: string): TXlsItemObj;
var
  Book: TXlsItemObj;
  Opened: IDispatch;
begin
  Result := nil;

  Opened := IWorkbooks.Open( AFileName,
      EmptyParam, False, EmptyParam, EmptyParam,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam,
      EmptyParam, EmptyParam, EmptyParam, EmptyParam, DEF_LCID );

  Book := CreateItem( TXlsWorkbook, Opened );

  Assert( Book <> nil );

  Add( Book );

  Result := Book;
end;

{ TXlsItemObj }

constructor TXlsItemObj.Create;
begin
  FItemDisp := nil;
end;

destructor TXlsItemObj.Destroy;
begin
  FItemDisp := nil;
  inherited;
end;

procedure TXlsItemObj.SetItemDisp(const Value: IDispatch);
begin
  FItemDisp := Value;
end;

{ TXlsWorksheet }

constructor TXlsWorksheet.Create;
begin
  inherited;
  FRanges := nil;
  FRanges := TXlsRanges.Create( nil );
end;

function TXlsWorksheet.GetAsWorksheet: IXlsSheet;
begin
  Assert( ItemDisp <> nil );

  Result := ItemDisp as IXlsSheet;
end;


function TXlsWorksheet.GetSheetRange(AFrom, ATo: string): IXlsRange;
begin
  Result := nil;

  if ATo <> '' then
    Result := GetRange( IWorksheet, AFrom, ATo )
  else
    Result := GetRange( IWorksheet, AFrom );

end;

function TXlsWorksheet.GetSheetRange(ARow, ACol: integer): IXlsRange;
begin
  Result := GetRange( IWorksheet, ARow, ACol );
end;

function TXlsWorksheet.GetCompareName: string;
begin
  Result := IWorksheet.Name;
end;

function TXlsWorksheet.GetSheetRange(ARngName: string): IXlsRange;
begin
  Result := GetRange( IWorksheet, ARngName );
end;

procedure TXlsWorksheet.SetItemDisp(const Value: IDispatch);
begin
  inherited SetItemDisp( Value );

  Assert( FRanges <> nil );

  if Value <> nil then
  begin
    FRanges.ContnrDisp := Value;
  end
  else
  begin
    FRanges.Clear;
  end;
end;

procedure TXlsWorksheet.Delete;
var
  Old: boolean;
begin
  IWorksheet.Delete( DEF_LCID );
end;

destructor TXlsWorksheet.Destroy;
begin
  FRanges.Free;
  FRanges := nil;

  inherited;
end;

{ TXlsWorksheets }

function TXlsWorksheets.AddItem: TXlsItemObj;
var
  AfterSheet: IXlsSheet;
  NewSheet: IDispatch;
begin
  Result := nil;

  AfterSheet := IWorksheets.Item[ IWorksheets.Count ] as IXlsSheet;
  NewSheet   := IWorksheets.Add( EmptyParam,
                  AfterSheet, DEF_ADD_SHEETS, xlWorksheet, DEF_LCID );

  Result := CreateItem( TXlsWorksheet, NewSheet );

  Assert( Result <> nil );

  Add( Result );
end;

{
procedure TXlsWorksheets.DoDeleteItem(AItem: TObject);
begin
  TXlsWorksheet(AItem).IWorksheet.Delete( DEF_LCID );
end;
}

function TXlsWorksheets.GetAsIWoorksheet(AName: string): IXlsSheet;
var
  Sh: TXlsItemObj;
begin
  Result := nil;

  Sh := nil;
  Sh := GetByName( AName );

  if Sh <> nil then
    Result := TXlsWorksheet(sh).IWorksheet;
end;

function TXlsWorksheets.GetAsWorksheets: IXlsSheets;
begin
  Assert( ContnrDisp <> nil );

  Result := ContnrDisp as IXlsSheets;
end;

{ TXlsRange }

function TXlsRange.GetAsRange: IXlsRange;
begin
  Assert( ItemDisp <> nil );

  Result := ItemDisp as IXlsRange;
end;

function TXlsRange.GetCompareName: string;
begin
  Result := '';
  try
    Result := IRange.Name.Name;
  except
    //Range have no name then try get address name
    try
      Result := IRange.Address[ False, False, xlA1, EmptyParam, EmptyParam ];
    except
      //may be need to re-raise or remove from try except
    end;
  end;
end;

{ TXlsRanges }

function TXlsRanges.Add(ARngName, AFrom, ATo: string): TXlsRange;
var
  RngTemp: IXlsRange;
begin
  Result := nil;

  RngTemp := nil;
  RngTemp := GetRange( (ContnrDisp as IXlsSheet), AFrom, ATo );
  if RngTemp <> nil then
  begin
    RngTemp.Name := ARngName;

    Result := TXlsRange(CreateItem( TXlsRange, RngTemp ));

    Assert( Result <> nil );

    Add( Result );
  end;
end;

function TXlsRanges.Add(ARngName: string; ARow, ACol: integer): TXlsRange;
var
  RngTemp: IXlsRange;
begin
  Result := nil;

  RngTemp := nil;
  RngTemp := GetRange( (ContnrDisp as IXlsSheet), ARow, ACol );
  if RngTemp <> nil then
  begin
    RngTemp.Name := ARngName;

    Result := TXlsRange(CreateItem( TXlsRange, RngTemp ));

    Assert( Result <> nil );

    Add( Result );
  end;
end;

function TXlsRanges.Add(ARngName, AAdrress: string): TXlsRange;
var
  RngTemp: IXlsRange;
begin
  Result := nil;

  RngTemp := nil;
  RngTemp := (ContnrDisp as IXlsSheet).Range[ AAdrress, EmptyParam ];
  if RngTemp <> nil then
  begin
    RngTemp.Name := ARngName;

    Result := TXlsRange(CreateItem( TXlsRange, RngTemp ));

    Assert( Result <> nil );

    Add( Result );
  end;
end;

function TXlsRanges.AddItem: TXlsItemObj;
begin
  Result := nil;
end;

function TXlsRanges.FindAsIRange(AName: string): IXlsRange;
var
  FndItm: TXlsRange;
begin
  Result := nil;

  FndItm := nil;
  FndItm := FindRange( AName );
  if FndItm <> nil then
  begin
    Result := FndItm.IRange;
  end;
end;

function TXlsRanges.FindRange(AName: string): TXlsRange;
var
  FndItm: TXlsItemObj;
begin
  Result := nil;

  //Seek in the container
  FndItm := GetByName( AName );
  //if not find seek in the excel;
  if FndItm = nil then
  begin
    FndItm := GetByNameFromSheet( AName );
  end;

  Result := TXlsRange(FndItm);
end;

function TXlsRanges.GetByNameFromSheet(AName: string): TXlsRange;
var
  RngTemp: IXlsRange;
begin
  Result := nil;

  RngTemp := nil;
  RngTemp := GetRange( (ContnrDisp as IXlsSheet), AName );
  if RngTemp <> nil then
  begin
    Result := TXlsRange(CreateItem( TXlsRange, RngTemp ));

    Assert( Result <> nil );

    Add( Result );
  end;
end;

end.
