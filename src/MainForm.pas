unit MainForm;

interface

uses
  Winapi.Windows,
  Winapi.ShellAPI,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  System.DateUtils,
  System.IOUtils,
  System.StrUtils,
  System.IniFiles,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.ExtCtrls,
  Vcl.ComCtrls,
  Vcl.Menus,
  Data.DB,
  FireDAC.Stan.Intf,
  FireDAC.Stan.Option,
  FireDAC.Stan.Error,
  FireDAC.UI.Intf,
  FireDAC.Phys.Intf,
  FireDAC.Stan.Def,
  FireDAC.Stan.Pool,
  FireDAC.Stan.Async,
  FireDAC.Phys,
  FireDAC.Phys.MSAcc,
  FireDAC.Phys.MSAccDef,
  FireDAC.Phys.ODBCBase,
  FireDAC.VCLUI.Wait,
  FireDAC.Comp.Client,
  FireDAC.Comp.DataSet,
  FireDAC.DApt,
  Vcl.Printers,
  Winapi.ActiveX,
  System.Win.Registry,
  SHDocVw;

type
  TfrmMain = class(TForm)
    pnlTop: TPanel;
    lblFile: TLabel;
    edtFile: TEdit;
    btnBrowse: TButton;
    lblFrom: TLabel;
    dtpFrom: TDateTimePicker;
    lblTo: TLabel;
    dtpTo: TDateTimePicker;
    btnRange: TButton;
    btnGenerate: TButton;
    btnSave: TButton;
    btnOpenBrowser: TButton;
    pnlStatus: TPanel;
    memHTML: TMemo;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    popRange: TPopupMenu;
    miToday: TMenuItem;
    miYesterday: TMenuItem;
    miSep1: TMenuItem;
    miCurMonth: TMenuItem;
    miPrevMonth: TMenuItem;
    miCurQuarter: TMenuItem;
    miPrevQuarter: TMenuItem;
    miCurYear: TMenuItem;
    miPrevYear: TMenuItem;
    miSep2: TMenuItem;
    miLast30: TMenuItem;
    miLast90: TMenuItem;
    miLast120: TMenuItem;
    FDConn: TFDConnection;
    lblReportType: TLabel;
    cmbReportType: TComboBox;
    btnSavePdf: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnBrowseClick(Sender: TObject);
    procedure btnRangeClick(Sender: TObject);
    procedure RangeMenuClick(Sender: TObject);
    procedure btnGenerateClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnOpenBrowserClick(Sender: TObject);
    procedure btnSavePdfClick(Sender: TObject);
  private
    FLastHTML: string;
    // Ukryta przeglądarka używana wyłącznie do "cichego" drukowania raportu
    // do PDF (przez wirtualną drukarkę "Microsoft Print to PDF"). Tworzona
    // w kodzie (nie w .dfm), żeby uniknąć ręcznego edytowania binarnych
    // danych ActiveX w pliku formularza.
    FWebBrowser: TWebBrowser;
    // Dane firmy - wczytane z tabeli 'Firma' przy połączeniu z bazą,
    // wykorzystywane w nagłówku obu raportów.
    FFirmaNazwa, FFirmaUlica, FFirmaNrDomu, FFirmaNrLokalu,
      FFirmaKodPoczt, FFirmaMiejscowosc, FFirmaNIP: string;
    FFirmaZaladowana: Boolean;
    function IniFileName: string;
    procedure LoadSettings;
    procedure SaveSettings;
    procedure ConnectToDatabase(const AFileName: string);
    procedure LoadFirmaData;
    procedure SetRange(ADateFrom, ADateTo: TDateTime);
    function BuildReportHTML(ADateFrom, ADateTo: TDateTime): string;
    function BuildPurchaseReportHTML(ADateFrom, ADateTo: TDateTime): string;
    function BuildSalesReportHTML(ADateFrom, ADateTo: TDateTime): string;
    function ReportFileTag: string;
    function CompanyHeaderHTML: string;
    procedure SetStatus(const AMsg: string);
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

type
  // Bufor pojedynczej pozycji dokumentu (do klasyfikacji całego dokumentu
  // na podstawie pola 'kosztowy' przed wyrenderowaniem nagłówka).
  TPozRow = record
    Nazwa, Jm: string;
    Ilosc, CenaN, CenaB: Variant;
    Kosztowy: Boolean;
  end;

const
  // typdok.id dla FZ i KFZ - ustalone z analizy bazy
  TYP_FZ  = 20;
  TYP_KFZ = 21;

{ ============================ helpers ============================ }

function HtmlEscape(const S: string): string;
var
  I: Integer;
  C: Char;
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create(Length(S));
  try
    for I := 1 to Length(S) do
    begin
      C := S[I];
      case C of
        '&': SB.Append('&amp;');
        '<': SB.Append('&lt;');
        '>': SB.Append('&gt;');
        '"': SB.Append('&quot;');
        '''': SB.Append('&#39;');
      else
        SB.Append(C);
      end;
    end;
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function FmtMoney(const V: Variant): string;
var
  D: Double;
begin
  if VarIsNull(V) or VarIsClear(V) then
    D := 0
  else
    D := V;
  Result := FormatFloat('#,##0.00', D);
end;

function FmtQty(const V: Variant): string;
var
  D: Double;
begin
  if VarIsNull(V) or VarIsClear(V) then
    D := 0
  else
    D := V;
  // ilości pokazujemy z 3 miejscami, ale obcinamy końcowe zera
  Result := FormatFloat('0.###', D);
end;

function VarToStrSafe(const V: Variant): string;
begin
  if VarIsNull(V) or VarIsClear(V) then
    Result := ''
  else
    Result := VarToStr(V);
end;

{ Bezpieczny odczyt pola jako Boolean - obsługuje zarówno BOOLEAN jak i numeryczne
  reprezentacje (Integer/SmallInt/Byte). Pole NULL traktujemy jako False. }
function FieldAsBoolSafe(F: TField): Boolean;
var
  V: Variant;
begin
  Result := False;
  if (F = nil) or F.IsNull then Exit;
  V := F.Value;
  if VarIsType(V, varBoolean) then
    Result := Boolean(V)
  else if VarIsNumeric(V) then
    Result := (Double(V) <> 0);
end;

{ Pierwszy/ostatni dzień kwartału obejmującego ADate }
procedure QuarterBounds(ADate: TDateTime; out AStart, AEnd: TDateTime);
var
  Y, M, D: Word;
  QStartMonth: Word;
begin
  DecodeDate(ADate, Y, M, D);
  QStartMonth := ((M - 1) div 3) * 3 + 1;
  AStart := EncodeDate(Y, QStartMonth, 1);
  AEnd   := EndOfTheMonth(EncodeDate(Y, QStartMonth + 2, 1));
end;

{ ============================ TfrmMain ============================ }

procedure TfrmMain.FormCreate(Sender: TObject);
var
  Today: TDateTime;
  DefaultMdb: string;
begin
  Today := Date;
  // Domyślnie - bieżący miesiąc
  SetRange(StartOfTheMonth(Today), EndOfTheMonth(Today));

  // Domyślny typ raportu - zakupy (zachowuje dotychczasowe zachowanie programu).
  cmbReportType.ItemIndex := 0;

  // Ukryta przeglądarka do drukowania raportu do PDF. Tworzona w kodzie,
  // z rodzicem ustawionym na formularz, ale niewidoczna i poza obszarem
  // roboczym - nie jest częścią normalnego interfejsu.
  FWebBrowser := TWebBrowser.Create(Self);
  // TWebBrowser (jako kontrolka ActiveX/TOleControl) nie pozwala na
  // bezpośrednie przypisanie "Parent := ..." przy tworzeniu w runtime
  // (błąd kompilatora E2129 "Cannot assign to a read-only property").
  // InsertControl robi dokładnie to samo, czego normalnie użyłoby
  // przypisanie Parent, tylko bez wywoływania tego konkretnego settera.
  InsertControl(FWebBrowser);
  FWebBrowser.Visible := False;
  FWebBrowser.Left := -3000;
  FWebBrowser.Top := -3000;
  // WAŻNE: silnik Trident/MSHTML liczy CSS-owe szerokości procentowe
  // (nasza tabela ma "width: 100%") względem rozmiaru samej kontrolki
  // WebBrowser, NIE względem szerokości strony drukarki - nawet przy
  // druku. Wcześniejszy rozmiar 100x100 px sprawiał, że cała tabela była
  // fizycznie wyliczona (a potem wydrukowana) jako 100 px szeroka, więc
  // większość kolumn i tak wychodziła poza tę wąską szerokość i była
  // ucinana. Ustawiamy więc rozmiar zbliżony do RZECZYWISTEGO obszaru
  // drukowalnego strony A4 (przy typowym 96 dpi i marginesach 0.4",
  // wymuszanych w btnSavePdfClick: (8.27" - 2*0.4") * 96 ≈ 750 px),
  // żeby układ tabeli był liczony względem realistycznej szerokości.
  FWebBrowser.Width := 750;
  FWebBrowser.Height := 1060;

  // Krok 1: spróbuj odtworzyć ścieżkę z poprzedniego uruchomienia z pliku INI.
  LoadSettings;

  // Krok 2: jeśli INI nic nie dało, a obok exe leży fakturka.mdb - podpowiedz ją.
  if Trim(edtFile.Text) = '' then
  begin
    DefaultMdb := TPath.Combine(ExtractFilePath(ParamStr(0)), 'fakturka.mdb');
    if FileExists(DefaultMdb) then
      edtFile.Text := DefaultMdb;
  end;

  SetStatus('Gotowy. Plik ustawień: ' + IniFileName + '. Wskaż plik bazy i kliknij "Generuj raport".');
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveSettings;
end;

function TfrmMain.IniFileName: string;
begin
  // Jeśli program został uruchomiony z parametrem - traktujemy go jako ścieżkę
  // do pliku INI, który ma być użyty (np. do obsługi kilku różnych baz/firm
  // z jednego programu, każda z własną konfiguracją).
  // Bez parametru - używamy domyślnego pliku MFRaporty.ini obok exe.
  if ParamCount >= 1 then
    Result := ParamStr(1)
  else
    Result := TPath.Combine(ExtractFilePath(ParamStr(0)), 'MFRaporty.ini');
end;

procedure TfrmMain.LoadSettings;
var
  Ini: TIniFile;
  Path: string;
begin
  if not FileExists(IniFileName) then
    Exit;
  Ini := TIniFile.Create(IniFileName);
  try
    // Ostatnio używany plik bazy - wpisujemy do pola tylko jeśli nadal istnieje.
    // Gdyby plik został przeniesiony lub usunięty, wolimy zostawić puste pole,
    // niż zaczynać od ścieżki która i tak rzuci błąd.
    Path := Ini.ReadString('Database', 'LastFile', '');
    if (Path <> '') and FileExists(Path) then
      edtFile.Text := Path;

    // Pamiętamy też ostatnie foldery dialogów - drobny komfort, ale dużo wart
    // gdy plik INI ma ustawioną ostatnią ścieżkę dla otwieracza i zapisywacza.
    Path := Ini.ReadString('Dialogs', 'OpenInitialDir', '');
    if (Path <> '') and DirectoryExists(Path) then
      dlgOpen.InitialDir := Path;

    Path := Ini.ReadString('Dialogs', 'SaveInitialDir', '');
    if (Path <> '') and DirectoryExists(Path) then
      dlgSave.InitialDir := Path;

    // Ostatnio wybrany typ raportu (0 = zakupy, 1 = sprzedaż).
    cmbReportType.ItemIndex := Ini.ReadInteger('Report', 'Type', cmbReportType.ItemIndex);
    if (cmbReportType.ItemIndex < 0) or (cmbReportType.ItemIndex >= cmbReportType.Items.Count) then
      cmbReportType.ItemIndex := 0;
  finally
    Ini.Free;
  end;
end;

procedure TfrmMain.SaveSettings;
var
  Ini: TIniFile;
begin
  try
    Ini := TIniFile.Create(IniFileName);
    try
      Ini.WriteString('Database', 'LastFile', Trim(edtFile.Text));
      // InitialDir aktualizujemy w klikach, ale jeszcze raz na zamknięciu - na wypadek
      // gdyby ostatnio użyty folder był inny niż aktualnie ustawiony w dialogach.
      Ini.WriteString('Dialogs', 'OpenInitialDir', dlgOpen.InitialDir);
      Ini.WriteString('Dialogs', 'SaveInitialDir', dlgSave.InitialDir);
      Ini.WriteInteger('Report', 'Type', cmbReportType.ItemIndex);
    finally
      Ini.Free;
    end;
  except
    // Cicho - jeżeli INI nie da się zapisać (np. read-only katalog programu),
    // nie chcemy z tego powodu psuć zamknięcia aplikacji.
  end;
end;

procedure TfrmMain.SetStatus(const AMsg: string);
begin
  pnlStatus.Caption := '   ' + AMsg;
  pnlStatus.Update;
end;

procedure TfrmMain.btnBrowseClick(Sender: TObject);
begin
  if edtFile.Text <> '' then
    dlgOpen.FileName := edtFile.Text;
  if dlgOpen.Execute then
  begin
    edtFile.Text := dlgOpen.FileName;
    // Zapamiętaj folder do następnego razu (zarówno w runtime jak i w INI).
    dlgOpen.InitialDir := ExtractFilePath(dlgOpen.FileName);
    SaveSettings;
  end;
end;

procedure TfrmMain.btnRangeClick(Sender: TObject);
var
  P: TPoint;
begin
  P := btnRange.ClientToScreen(Point(0, btnRange.Height));
  popRange.Popup(P.X, P.Y);
end;

procedure TfrmMain.SetRange(ADateFrom, ADateTo: TDateTime);
begin
  dtpFrom.Date := ADateFrom;
  dtpTo.Date   := ADateTo;
end;

procedure TfrmMain.RangeMenuClick(Sender: TObject);
var
  Today, D1, D2: TDateTime;
  Y, M, Dummy: Word;
begin
  Today := Date;
  D1 := Today;
  D2 := Today;
  case TMenuItem(Sender).Tag of
    1: // Dzień dzisiejszy
       begin D1 := Today; D2 := Today; end;
    2: // Dzień wczorajszy
       begin D1 := Today - 1; D2 := Today - 1; end;
    3: // Poprzedni miesiąc
       begin
         D1 := StartOfTheMonth(IncMonth(Today, -1));
         D2 := EndOfTheMonth(IncMonth(Today, -1));
       end;
    4: // Poprzedni kwartał
       begin
         QuarterBounds(IncMonth(Today, -3), D1, D2);
       end;
    5: // Poprzedni rok
       begin
         DecodeDate(Today, Y, M, Dummy);
         D1 := EncodeDate(Y - 1, 1, 1);
         D2 := EncodeDate(Y - 1, 12, 31);
       end;
    6: // Bieżący miesiąc
       begin
         D1 := StartOfTheMonth(Today);
         D2 := EndOfTheMonth(Today);
       end;
    7: // Bieżący kwartał
       begin
         QuarterBounds(Today, D1, D2);
       end;
    8: // Bieżący rok
       begin
         DecodeDate(Today, Y, M, Dummy);
         D1 := EncodeDate(Y, 1, 1);
         D2 := EncodeDate(Y, 12, 31);
       end;
    9: // Ostatnie 30 dni (wliczając dzisiaj)
       begin D1 := Today - 29; D2 := Today; end;
   10: // Ostatnie 90 dni
       begin D1 := Today - 89; D2 := Today; end;
   11: // Ostatnie 120 dni
       begin D1 := Today - 119; D2 := Today; end;
  end;
  SetRange(D1, D2);
end;

procedure TfrmMain.ConnectToDatabase(const AFileName: string);
var
  IsACE: Boolean;
  FS: TFileStream;
  Buf: array[0..19] of Byte;
begin
  if not FileExists(AFileName) then
    raise Exception.CreateFmt('Plik bazy nie istnieje: %s', [AFileName]);

  if FDConn.Connected then
    FDConn.Close;

  // Sprawdzamy SYGNATURĘ pliku - mimo rozszerzenia .mdb plik może być fizycznie
  // w formacie ACE (Access 2007+). W nagłówku na offsetcie 4 mamy tekst
  // "Standard Jet DB" lub "Standard ACE DB".
  IsACE := False;
  try
    FS := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
    try
      if FS.Size > SizeOf(Buf) then
      begin
        FS.ReadBuffer(Buf, SizeOf(Buf));
        // 'A' na poz. 13, 'C' na poz. 14 -> ACE; 'J' na poz. 13 -> Jet
        IsACE := (Buf[13] = Ord('A')) and (Buf[14] = Ord('C'));
      end;
    finally
      FS.Free;
    end;
  except
    // jak nie odczytamy - polecimy domyślnie przez MSAcc
  end;

  FDConn.Params.Clear;
  FDConn.Params.Add('DriverID=MSAcc');
  FDConn.Params.Add('Database=' + AFileName);
  if IsACE then
    // Wymusza użycie sterownika ACE OLE DB - kluczowe, gdy plik .mdb jest
    // fizycznie w formacie ACE (Access 2007+). Sterownik musi być zainstalowany
    // w wersji zgodnej z bitnością aplikacji (32 lub 64 bit).
    FDConn.Params.Add('Provider=Microsoft.ACE.OLEDB.12.0');
  // Bez hasła. Jeśli baza jest chroniona - dopisz: Password=...

  try
    FDConn.Open;
  except
    on E: Exception do
      raise Exception.CreateFmt(
        'Nie udało się otworzyć bazy %s'#13#10 +
        'Wykryty format pliku: %s'#13#10 +
        'Błąd: %s'#13#10#13#10 +
        'Najczęstsza przyczyna: brak Microsoft Access Database Engine ' +
        'w wersji zgodnej z bitnością aplikacji (32/64). ' +
        'Pobierz instalator ze strony Microsoftu: AccessDatabaseEngine.exe ' +
        'dla aplikacji 32-bit lub AccessDatabaseEngine_X64.exe dla 64-bit.',
        [AFileName,
         IfThen(IsACE, 'ACE (Access 2007+)', 'Jet (Access 2000-2003)'),
         E.Message]);
  end;

  // Wczytaj dane firmy z tabeli 'Firma' - zostaną wykorzystane w nagłówku
  // obu raportów. Robimy to raz, zaraz po połączeniu, żeby nie odpytywać
  // bazy przy każdym generowaniu raportu.
  LoadFirmaData;
end;

procedure TfrmMain.LoadFirmaData;
var
  q: TFDQuery;
begin
  // Resetujemy dane przy każdym (re)połączeniu - żeby dane z poprzednio
  // otwartej bazy nie "zostały" w nagłówku, gdyby nowa baza nie miała
  // tabeli 'Firma' albo była pusta.
  FFirmaNazwa       := '';
  FFirmaUlica       := '';
  FFirmaNrDomu      := '';
  FFirmaNrLokalu    := '';
  FFirmaKodPoczt    := '';
  FFirmaMiejscowosc := '';
  FFirmaNIP         := '';
  FFirmaZaladowana  := False;

  q := TFDQuery.Create(nil);
  try
    q.Connection := FDConn;
    try
      q.SQL.Text :=
        'SELECT Nazwa, ulica, nrdomu, nrlokalu, kodpoczt, miejscowosc, NIP ' +
        'FROM Firma';
      q.Open;
      if not q.Eof then
      begin
        FFirmaNazwa       := VarToStrSafe(q.FieldByName('Nazwa').Value);
        FFirmaUlica       := VarToStrSafe(q.FieldByName('ulica').Value);
        FFirmaNrDomu      := VarToStrSafe(q.FieldByName('nrdomu').Value);
        FFirmaNrLokalu    := VarToStrSafe(q.FieldByName('nrlokalu').Value);
        FFirmaKodPoczt    := VarToStrSafe(q.FieldByName('kodpoczt').Value);
        FFirmaMiejscowosc := VarToStrSafe(q.FieldByName('miejscowosc').Value);
        FFirmaNIP         := VarToStrSafe(q.FieldByName('NIP').Value);
        FFirmaZaladowana  := True;
      end;
      q.Close;
    except
      // Jeśli tabeli 'Firma' nie ma (albo ma inną strukturę niż oczekiwana) -
      // nie przerywamy działania programu. Nagłówek raportu po prostu nie
      // będzie zawierał danych firmy.
    end;
  finally
    q.Free;
  end;
end;

function TfrmMain.BuildPurchaseReportHTML(ADateFrom, ADateTo: TDateTime): string;
const
  // Pobieramy dokumenty zakupowe wraz z pozycjami w jednym zapytaniu
  // (nagłówek wielokrotnie powtarzany jako lewa strona LEFT JOIN-a).
  // Filtr: typ in (FZ, KFZ), zakres dat wystawienia, status<>1 (nie skasowane).
  // Sortowanie: data, numer dokumentu, kolejność pozycji.
  SQL_DOK =
    'SELECT '+
    '  d.id              AS d_id, '+
    '  d.nazwa           AS d_nazwa, '+
    '  td.skrot          AS d_skrot, '+
    '  d.wyst            AS d_wyst, '+
    '  d.sprz            AS d_sprz, '+
    '  d.nazwak          AS k_nazwa, '+
    '  d.nip             AS k_nip, '+
    '  d.ulica           AS k_ulica, '+
    '  d.nrdomu          AS k_nrdomu, '+
    '  d.nrlokalu        AS k_nrlokalu, '+
    '  d.kod             AS k_kod, '+
    '  d.miasto          AS k_miasto, '+
    '  d.netto           AS d_netto, '+
    '  d.vat             AS d_vat, '+
    '  d.brutto          AS d_brutto, '+
    '  d.waluta          AS d_waluta, '+
    '  d.platnosc        AS d_platnosc, '+
    '  d.termin          AS d_termin, '+
    '  d.bezterminu      AS d_bezterminu, '+
    '  d.podzielnaplatnosc AS d_mpp, '+
    '  d.zaplacony       AS d_zaplacony '+
    'FROM dok d LEFT JOIN typdok td ON td.id = d.typ '+
    'WHERE d.typ IN (:t1, :t2) '+
    '  AND (d.status IS NULL OR d.status <> 1) '+
    '  AND d.wyst >= :df AND d.wyst < :dt '+
    'ORDER BY d.wyst, d.nazwa';

  SQL_POZ =
    'SELECT '+
    '  p.id          AS p_id, '+
    '  p.nazwa       AS p_nazwa, '+
    '  p.ilosc       AS p_ilosc, '+
    '  p.jm          AS p_jm, '+
    '  p.cena        AS p_cena_n, '+
    '  p.cenab       AS p_cena_b, '+
    '  p.vat         AS p_vat, '+
    '  p.netto       AS p_netto, '+
    '  p.brutto      AS p_brutto, '+
    '  p.kosztowy    AS p_koszt '+
    'FROM dane p '+
    'WHERE p.nrdok = :idd '+
    'ORDER BY p.id';
var
  qDok, qPoz: TFDQuery;
  SB: TStringBuilder;
  TotalNetto, TotalVat, TotalBrutto: Double;
  TotalNettoK, TotalVatK, TotalBruttoK: Double;
  TotalNettoS, TotalVatS, TotalBruttoS: Double;
  CntDok, CntPoz, CntK, CntS: Integer;
  Adres, Waluta, NrDok: string;
  // Bufor pozycji jednego dokumentu - musimy najpierw zebrać wszystkie pozycje,
  // żeby ustalić klasę całego dokumentu (K jeśli choć jedna pozycja kosztowa).
  Pozycje: TArray<TPozRow>;
  DocIsKoszt: Boolean;
  KlasaCSS, KlasaLit, KlasaOpis: string;
  PR: TPozRow;
  Lp, I: Integer;
  DN, DV, DB: Double;
  PlatnoscOpis, ZaplacOpis: string;
  ZaplaconaFakt, BezTerm: Boolean;
  TerminDt: TDateTime;
begin
  qDok := TFDQuery.Create(nil);
  qPoz := TFDQuery.Create(nil);
  SB   := TStringBuilder.Create;
  try
    qDok.Connection := FDConn;
    qPoz.Connection := FDConn;

    qDok.SQL.Text := SQL_DOK;
    qDok.ParamByName('t1').AsInteger := TYP_FZ;
    qDok.ParamByName('t2').AsInteger := TYP_KFZ;
    qDok.ParamByName('df').AsDateTime := DateOf(ADateFrom);
    // Górna granica jako "< (data_do + 1 dzień)" - bezpieczne dla dat z czasem
    qDok.ParamByName('dt').AsDateTime := DateOf(ADateTo) + 1;
    qDok.Open;

    qPoz.SQL.Text := SQL_POZ;

    TotalNetto  := 0;  TotalVat  := 0;  TotalBrutto  := 0;
    TotalNettoK := 0;  TotalVatK := 0;  TotalBruttoK := 0;
    TotalNettoS := 0;  TotalVatS := 0;  TotalBruttoS := 0;
    CntDok := 0;  CntPoz := 0;  CntK := 0;  CntS := 0;

    // ===== HTML =====
    SB.AppendLine('<!DOCTYPE html>');
    SB.AppendLine('<html lang="pl">');
    SB.AppendLine('<head>');
    SB.AppendLine('<meta charset="UTF-8">');
    SB.AppendLine('<title>Raport zakupów ' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateFrom)) + ' - ' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateTo)) + '</title>');
    SB.AppendLine('<style>');
    SB.AppendLine('  body { font-family: "Segoe UI", Arial, sans-serif; font-size: 12px; color: #222; margin: 20px; }');
    SB.AppendLine('  h1 { font-size: 18px; margin: 0 0 4px 0; }');
    SB.AppendLine('  .meta { color: #666; margin-bottom: 16px; }');
    // table-layout: fixed + word-wrap - bez tego długi tekst (np. nazwa
    // towaru albo kontrahenta) potrafi rozciągnąć tabelę szerzej niż
    // strona, a wtedy druk/PDF przycina wszystko poza prawą krawędzią
    // zamiast zawijać tekst do kolejnej linii.
    SB.AppendLine('  table { border-collapse: collapse; width: 100%; max-width: 100%; table-layout: fixed; margin-bottom: 18px; }');
    SB.AppendLine('  th, td { border: 1px solid #d0d0d0; padding: 4px 6px; vertical-align: top; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; }');
    SB.AppendLine('  th { background: #f0f0f0; text-align: left; font-weight: 600; }');
    SB.AppendLine('  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }');
    SB.AppendLine('  td.ctr, th.ctr { text-align: center; }');
    // Nagłówek dokumentu - kolor zależy od klasyfikacji K (czerwony) / S (zielony).
    SB.AppendLine('  .doc-head td { font-weight: 600; }');
    SB.AppendLine('  .doc-k .doc-head { background: #fde2e2; }');
    SB.AppendLine('  .doc-s .doc-head { background: #dff5d8; }');
    // Lewa krawędź tabeli wyróżniająca klasę dokumentu.
    SB.AppendLine('  table.doc-k { border-left: 4px solid #c62828; }');
    SB.AppendLine('  table.doc-s { border-left: 4px solid #2e7d32; }');
    SB.AppendLine('  .doc-sum { background: #f7f7f7; }');
    SB.AppendLine('  .doc-sum td { font-weight: 600; }');
    SB.AppendLine('  .grand { background: #fff4d6; }');
    SB.AppendLine('  .grand td { font-weight: 700; font-size: 13px; }');
    SB.AppendLine('  .pos-name { width: 50%; }');
    // Znaczek K/S - duża, czytelna plakietka po prawej stronie nagłówka.
    SB.AppendLine('  .badge { float: right; display: inline-block; min-width: 26px; padding: 2px 10px;');
    SB.AppendLine('           border-radius: 4px; color: #fff; font-weight: 700; font-size: 14px;');
    SB.AppendLine('           text-align: center; letter-spacing: 1px; }');
    SB.AppendLine('  .badge-k { background: #c62828; }');
    SB.AppendLine('  .badge-s { background: #2e7d32; }');
    // Wiersz pozycji oznaczony jako kosztowy - delikatny czerwonawy odcień.
    SB.AppendLine('  tr.poz-k td { background: #fff5f5; }');
    SB.AppendLine('  .legend { margin-top: 6px; color: #555; font-size: 11px; }');
    SB.AppendLine('  .empty { color: #999; font-style: italic; padding: 20px; text-align: center; }');
    // Pasek z informacjami o płatności pod adresem kontrahenta.
    SB.AppendLine('  .pay-row { margin-top: 4px; font-weight: 400; color: #333; }');
    SB.AppendLine('  .pay-row .lbl { color: #666; }');
    SB.AppendLine('  .pay-row .val { font-weight: 600; }');
    // Małe pigułki statusu (zapłacone / niezapłacone / przeterminowane / MPP).
    SB.AppendLine('  .pill { display: inline-block; padding: 1px 8px; border-radius: 10px;');
    SB.AppendLine('          font-size: 11px; font-weight: 600; margin-left: 6px; vertical-align: 1px; }');
    SB.AppendLine('  .pill-ok   { background: #dff5d8; color: #1b5e20; border: 1px solid #a5d6a7; }');
    SB.AppendLine('  .pill-warn { background: #ffe9b3; color: #8a5a00; border: 1px solid #f0c860; }');
    SB.AppendLine('  .pill-bad  { background: #fde2e2; color: #b71c1c; border: 1px solid #ef9a9a; }');
    SB.AppendLine('  .pill-info { background: #e3f2fd; color: #0d47a1; border: 1px solid #90caf9; }');
    SB.AppendLine('  .firma { margin-bottom: 14px; line-height: 1.4; }');
    SB.AppendLine('  .firma-nazwa { font-weight: 700; font-size: 14px; }');
    SB.AppendLine('  @media print { body { margin: 8mm; } .no-print { display: none; } }');
    SB.AppendLine('</style>');
    SB.AppendLine('</head>');
    SB.AppendLine('<body>');

    SB.Append(CompanyHeaderHTML);
    SB.AppendLine('<h1>Raport zakupów (FZ / KFZ)</h1>');
    SB.AppendLine('<div class="meta">Zakres dat wystawienia: <b>' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateFrom)) + '</b> &ndash; <b>' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateTo)) + '</b><br>' +
      'Wygenerowano: ' + HtmlEscape(FormatDateTime('yyyy-mm-dd hh:nn', Now)) + '</div>');

    if qDok.Eof then
    begin
      SB.AppendLine('<div class="empty">Brak dokumentów zakupowych w wybranym zakresie dat.</div>');
    end
    else
    begin
      while not qDok.Eof do
      begin
        Inc(CntDok);
        NrDok := VarToStrSafe(qDok.FieldByName('d_nazwa').Value);
        if NrDok = '' then
          NrDok := '(bez numeru, id=' + IntToStr(qDok.FieldByName('d_id').AsInteger) + ')';

        // Adres kontrahenta - sklejony ładnie
        Adres := '';
        if VarToStrSafe(qDok.FieldByName('k_ulica').Value) <> '' then
          Adres := Trim(VarToStrSafe(qDok.FieldByName('k_ulica').Value) + ' ' +
                        VarToStrSafe(qDok.FieldByName('k_nrdomu').Value));
        if VarToStrSafe(qDok.FieldByName('k_nrlokalu').Value) <> '' then
          Adres := Adres + '/' + VarToStrSafe(qDok.FieldByName('k_nrlokalu').Value);
        if (VarToStrSafe(qDok.FieldByName('k_kod').Value) <> '') or
           (VarToStrSafe(qDok.FieldByName('k_miasto').Value) <> '') then
        begin
          if Adres <> '' then Adres := Adres + ', ';
          Adres := Adres + Trim(VarToStrSafe(qDok.FieldByName('k_kod').Value) + ' ' +
                                VarToStrSafe(qDok.FieldByName('k_miasto').Value));
        end;

        Waluta := VarToStrSafe(qDok.FieldByName('d_waluta').Value);
        if Waluta = '' then Waluta := 'PLN';

        // === Informacje o płatności ===
        // Sposób płatności jest w 'platnosc' (tekst, np. "gotówka", "przelew 7", "karta").
        // Nie używamy słownika - pole jest już rozwinięte w bazie.
        PlatnoscOpis := Trim(VarToStrSafe(qDok.FieldByName('d_platnosc').Value));
        if PlatnoscOpis = '' then
          PlatnoscOpis := '(nie określono)';

        // Termin: jeśli flaga 'bezterminu' jest ustawiona - pomijamy datę
        // (program zapisuje wtedy sztuczną datę 01-01-0001).
        BezTerm := FieldAsBoolSafe(qDok.FieldByName('d_bezterminu'));
        TerminDt := 0;
        if not qDok.FieldByName('d_termin').IsNull then
          TerminDt := qDok.FieldByName('d_termin').AsDateTime;

        ZaplaconaFakt := FieldAsBoolSafe(qDok.FieldByName('d_zaplacony'));
        if ZaplaconaFakt then
          ZaplacOpis := '<span class="pill pill-ok">zapłacono</span>'
        else if (not BezTerm) and (TerminDt > 0) and (TerminDt < Date) then
          // Termin minął, faktura nadal niezapłacona - wyróżniamy mocniej.
          ZaplacOpis := '<span class="pill pill-bad">przeterminowane</span>'
        else
          ZaplacOpis := '<span class="pill pill-warn">do zapłaty</span>';

        // === Krok 1: wczytaj pozycje do bufora i wyznacz klasę dokumentu ===
        // Reguła klasyfikacji: jeśli choć jedna pozycja ma flagę 'kosztowy' = 1,
        // cały dokument oznaczamy jako KOSZTOWY (K). W przeciwnym razie - SPRZEDAŻOWY (S).
        SetLength(Pozycje, 0);
        DocIsKoszt := False;

        qPoz.Close;
        qPoz.ParamByName('idd').AsInteger := qDok.FieldByName('d_id').AsInteger;
        qPoz.Open;
        while not qPoz.Eof do
        begin
          PR.Nazwa    := VarToStrSafe(qPoz.FieldByName('p_nazwa').Value);
          PR.Ilosc    := qPoz.FieldByName('p_ilosc').Value;
          PR.Jm       := VarToStrSafe(qPoz.FieldByName('p_jm').Value);
          PR.CenaN    := qPoz.FieldByName('p_cena_n').Value;
          PR.CenaB    := qPoz.FieldByName('p_cena_b').Value;
          // Pole 'kosztowy' jest w bazie typu BOOLEAN. Czytamy bezpiecznie -
          // funkcja FieldAsBoolSafe obsługuje też numeryczne typy i NULL.
          PR.Kosztowy := FieldAsBoolSafe(qPoz.FieldByName('p_koszt'));
          if PR.Kosztowy then
            DocIsKoszt := True;
          SetLength(Pozycje, Length(Pozycje) + 1);
          Pozycje[High(Pozycje)] := PR;
          qPoz.Next;
        end;
        qPoz.Close;

        if DocIsKoszt then
        begin
          KlasaCSS  := 'doc-k';
          KlasaLit  := 'K';
          KlasaOpis := 'kosztowa';
          Inc(CntK);
        end
        else
        begin
          KlasaCSS  := 'doc-s';
          KlasaLit  := 'S';
          KlasaOpis := 'sprzedażowa';
          Inc(CntS);
        end;

        // === Krok 2: render tabeli dokumentu z odpowiednią klasą CSS ===
        SB.AppendLine('<table class="' + KlasaCSS + '">');
        // Nagłówek dokumentu - kontrahent + data + plakietka K/S po prawej.
        SB.AppendLine('<tr class="doc-head">');
        SB.AppendLine('  <td colspan="6">');
        SB.AppendLine('    <span class="badge badge-' + LowerCase(KlasaLit) + '" ' +
                       'title="Faktura ' + KlasaOpis + '">' + KlasaLit + '</span>');
        SB.AppendLine('    <b>' + HtmlEscape(qDok.FieldByName('d_skrot').AsString) +
                       ' &nbsp;' + HtmlEscape(NrDok) + '</b>');
        if not qDok.FieldByName('d_wyst').IsNull then
          SB.Append(' &nbsp;|&nbsp; data wystawienia: <b>' +
            FormatDateTime('yyyy-mm-dd', qDok.FieldByName('d_wyst').AsDateTime) + '</b>');
        SB.AppendLine('<br>');
        SB.AppendLine('    Kontrahent: <b>' +
          HtmlEscape(VarToStrSafe(qDok.FieldByName('k_nazwa').Value)) + '</b>');
        if VarToStrSafe(qDok.FieldByName('k_nip').Value) <> '' then
          SB.Append(' &nbsp;NIP: ' + HtmlEscape(VarToStrSafe(qDok.FieldByName('k_nip').Value)));
        if Adres <> '' then
          SB.Append('<br>    Adres: ' + HtmlEscape(Adres));

        // === Wiersz z informacjami o płatności ===
        // Sposób płatności + status zapłaty (z pigułką kolorową) + termin (tylko gdy ustalony).
        // MPP (mechanizm podzielonej płatności) - dodatkowa pigułka, jeśli ustawiona.
        // Uwaga: jeśli faktura jest "bez terminu" (flaga BezTerm=true), pomijamy pole
        // termin całkowicie - nie pokazujemy ani daty, ani placeholderowego napisu.
        SB.AppendLine('<br>');
        SB.Append('    <span class="pay-row"><span class="lbl">Płatność:</span> ' +
          '<span class="val">' + HtmlEscape(PlatnoscOpis) + '</span>' + ZaplacOpis);
        if (not BezTerm) and (TerminDt > 0) then
          SB.Append(' &nbsp;<span class="lbl">termin:</span> <span class="val">' +
            FormatDateTime('yyyy-mm-dd', TerminDt) + '</span>');
        if FieldAsBoolSafe(qDok.FieldByName('d_mpp')) then
          SB.Append('<span class="pill pill-info" title="Mechanizm podzielonej płatności">MPP</span>');
        SB.AppendLine('</span>');

        SB.AppendLine('  </td>');
        SB.AppendLine('</tr>');

        // Nagłówek tabeli pozycji
        SB.AppendLine('<tr>');
        SB.AppendLine('  <th class="ctr" style="width:32px">Lp.</th>');
        SB.AppendLine('  <th class="pos-name">Nazwa towaru / usługi</th>');
        SB.AppendLine('  <th class="num">Ilość</th>');
        SB.AppendLine('  <th class="ctr">J.m.</th>');
        SB.AppendLine('  <th class="num">Cena netto</th>');
        SB.AppendLine('  <th class="num">Cena brutto</th>');
        SB.AppendLine('</tr>');

        // Pozycje (z bufora)
        if Length(Pozycje) = 0 then
        begin
          SB.AppendLine('<tr><td colspan="6" class="empty">brak pozycji</td></tr>');
        end
        else
        begin
          for I := 0 to High(Pozycje) do
          begin
            Lp := I + 1;
            Inc(CntPoz);
            // Wiersz z pozycją kosztową dostaje delikatne tło, żeby od razu było widać,
            // która konkretnie pozycja zdecydowała o klasyfikacji "K".
            if Pozycje[I].Kosztowy then
              SB.AppendLine('<tr class="poz-k">')
            else
              SB.AppendLine('<tr>');
            SB.AppendLine('  <td class="ctr">' + IntToStr(Lp) + '</td>');
            SB.AppendLine('  <td>' + HtmlEscape(Pozycje[I].Nazwa) + '</td>');
            SB.AppendLine('  <td class="num">' + FmtQty(Pozycje[I].Ilosc) + '</td>');
            SB.AppendLine('  <td class="ctr">' + HtmlEscape(Pozycje[I].Jm) + '</td>');
            SB.AppendLine('  <td class="num">' + FmtMoney(Pozycje[I].CenaN) + '</td>');
            SB.AppendLine('  <td class="num">' + FmtMoney(Pozycje[I].CenaB) + '</td>');
            SB.AppendLine('</tr>');
          end;
        end;

        // Podsumowanie dokumentu
        SB.AppendLine('<tr class="doc-sum">');
        SB.AppendLine('  <td colspan="2" style="text-align:right">RAZEM dokument (' +
          HtmlEscape(Waluta) + '):</td>');
        SB.AppendLine('  <td class="num" colspan="2">netto: ' +
          FmtMoney(qDok.FieldByName('d_netto').Value) + '</td>');
        SB.AppendLine('  <td class="num">VAT: ' +
          FmtMoney(qDok.FieldByName('d_vat').Value) + '</td>');
        SB.AppendLine('  <td class="num">brutto: ' +
          FmtMoney(qDok.FieldByName('d_brutto').Value) + '</td>');
        SB.AppendLine('</tr>');
        SB.AppendLine('</table>');

        // Sumujemy tylko dokumenty w PLN do "wielkiego" podsumowania
        // (mieszanie walut nie miałoby sensu). Dodatkowo prowadzimy osobne sumy K i S.
        if SameText(Waluta, 'PLN') then
        begin
          DN := 0; DV := 0; DB := 0;
          if not qDok.FieldByName('d_netto').IsNull  then DN := qDok.FieldByName('d_netto').AsFloat;
          if not qDok.FieldByName('d_vat').IsNull    then DV := qDok.FieldByName('d_vat').AsFloat;
          if not qDok.FieldByName('d_brutto').IsNull then DB := qDok.FieldByName('d_brutto').AsFloat;
          TotalNetto  := TotalNetto  + DN;
          TotalVat    := TotalVat    + DV;
          TotalBrutto := TotalBrutto + DB;
          if DocIsKoszt then
          begin
            TotalNettoK  := TotalNettoK  + DN;
            TotalVatK    := TotalVatK    + DV;
            TotalBruttoK := TotalBruttoK + DB;
          end
          else
          begin
            TotalNettoS  := TotalNettoS  + DN;
            TotalVatS    := TotalVatS    + DV;
            TotalBruttoS := TotalBruttoS + DB;
          end;
        end;

        qDok.Next;
      end;

      // Podsumowanie zbiorcze - z rozbiciem na K i S oraz sumą całkowitą.
      SB.AppendLine('<table>');
      SB.AppendLine('<tr class="doc-head" style="background:#fde2e2">');
      SB.AppendLine('  <td><span class="badge badge-k">K</span> &nbsp; ' +
                    'Faktury kosztowe (PLN), dokumentów: ' + IntToStr(CntK) + '</td>');
      SB.AppendLine('  <td class="num">netto: '  + FmtMoney(TotalNettoK)  + '</td>');
      SB.AppendLine('  <td class="num">VAT: '    + FmtMoney(TotalVatK)    + '</td>');
      SB.AppendLine('  <td class="num">brutto: ' + FmtMoney(TotalBruttoK) + '</td>');
      SB.AppendLine('</tr>');
      SB.AppendLine('<tr class="doc-head" style="background:#dff5d8">');
      SB.AppendLine('  <td><span class="badge badge-s">S</span> &nbsp; ' +
                    'Faktury sprzedażowe (PLN), dokumentów: ' + IntToStr(CntS) + '</td>');
      SB.AppendLine('  <td class="num">netto: '  + FmtMoney(TotalNettoS)  + '</td>');
      SB.AppendLine('  <td class="num">VAT: '    + FmtMoney(TotalVatS)    + '</td>');
      SB.AppendLine('  <td class="num">brutto: ' + FmtMoney(TotalBruttoS) + '</td>');
      SB.AppendLine('</tr>');
      SB.AppendLine('<tr class="grand">');
      SB.AppendLine('  <td>RAZEM (PLN), dokumentów: ' + IntToStr(CntDok) +
                    ', pozycji: ' + IntToStr(CntPoz) + '</td>');
      SB.AppendLine('  <td class="num">netto: '  + FmtMoney(TotalNetto)  + '</td>');
      SB.AppendLine('  <td class="num">VAT: '    + FmtMoney(TotalVat)    + '</td>');
      SB.AppendLine('  <td class="num">brutto: ' + FmtMoney(TotalBrutto) + '</td>');
      SB.AppendLine('</tr>');
      SB.AppendLine('</table>');
      SB.AppendLine('<div class="legend">');
      SB.AppendLine('  <span class="badge badge-k">K</span> faktura kosztowa &nbsp;&ndash;&nbsp; ' +
                    'co najmniej jedna pozycja jest kosztem &nbsp;&nbsp;&nbsp; ');
      SB.AppendLine('  <span class="badge badge-s">S</span> faktura sprzedażowa &nbsp;&ndash;&nbsp; ' +
                    'żadna pozycja nie jest oznaczona jako koszt');
      SB.AppendLine('</div>');
      SB.AppendLine('<div class="meta">Uwaga: w sumie zbiorczej uwzględniono wyłącznie dokumenty w PLN. ' +
        'Dokumenty walutowe są widoczne na liście, ale nie sumują się do łącznej kwoty.</div>');
    end;

    SB.AppendLine('</body>');
    SB.AppendLine('</html>');

    Result := SB.ToString;
  finally
    qPoz.Free;
    qDok.Free;
    SB.Free;
  end;
end;

function TfrmMain.BuildSalesReportHTML(ADateFrom, ADateTo: TDateTime): string;
const
  // Dokumenty sprzedaży (FV) wraz z ich korektami (KFV). W odróżnieniu od
  // raportu zakupów (gdzie typ=20/21 zostały ustalone wprost z analizy
  // bazy), typ dokumentu jest tu rozpoznawany przez JOIN z 'typdok' i filtr
  // po skrócie - niezależnie od konkretnych numerów id w tej tabeli.
  // Proformy (FPV) są celowo pominięte - to dokumenty informacyjne, a nie
  // faktury sprzedaży w rozumieniu VAT.
  // Filtr: zakres dat wystawienia, status<>1 (nie skasowane).
  SQL_DOK_SPRZ =
    'SELECT '+
    '  d.id       AS d_id, '+
    '  d.nazwa    AS d_nazwa, '+
    '  d.wyst     AS d_wyst, '+
    '  d.nazwak   AS k_nazwa, '+
    '  d.platnosc AS d_platnosc, '+
    '  d.netto    AS d_netto, '+
    '  d.vat      AS d_vat, '+
    '  d.brutto   AS d_brutto '+
    'FROM dok d LEFT JOIN typdok td ON td.id = d.typ '+
    'WHERE td.skrot IN (''FV'', ''KFV'') '+
    '  AND (d.status IS NULL OR d.status <> 1) '+
    '  AND d.wyst >= :df AND d.wyst < :dt '+
    'ORDER BY d.wyst, d.nazwa';
var
  qDok: TFDQuery;
  SB: TStringBuilder;
  Lp: Integer;
  TotalNetto, TotalVat, TotalBrutto: Double;
  NrDok, Kontrahent, PlatnoscOpis: string;
begin
  qDok := TFDQuery.Create(nil);
  SB   := TStringBuilder.Create;
  try
    qDok.Connection := FDConn;
    qDok.SQL.Text := SQL_DOK_SPRZ;
    qDok.ParamByName('df').AsDateTime := DateOf(ADateFrom);
    // Górna granica jako "< (data_do + 1 dzień)" - bezpieczne dla dat z czasem.
    qDok.ParamByName('dt').AsDateTime := DateOf(ADateTo) + 1;
    qDok.Open;

    TotalNetto  := 0;
    TotalVat    := 0;
    TotalBrutto := 0;
    Lp := 0;

    // ===== HTML =====
    SB.AppendLine('<!DOCTYPE html>');
    SB.AppendLine('<html lang="pl">');
    SB.AppendLine('<head>');
    SB.AppendLine('<meta charset="UTF-8">');
    SB.AppendLine('<title>Raport sprzedaży ' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateFrom)) + ' - ' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateTo)) + '</title>');
    SB.AppendLine('<style>');
    SB.AppendLine('  body { font-family: "Segoe UI", Arial, sans-serif; font-size: 12px; color: #222; margin: 20px; }');
    SB.AppendLine('  h1 { font-size: 18px; margin: 0 0 4px 0; }');
    SB.AppendLine('  .meta { color: #666; margin-bottom: 16px; }');
    SB.AppendLine('  table { border-collapse: collapse; width: 100%; max-width: 100%; table-layout: fixed; margin-bottom: 18px; }');
    SB.AppendLine('  th, td { border: 1px solid #d0d0d0; padding: 4px 8px; vertical-align: top; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; }');
    SB.AppendLine('  th { background: #f0f0f0; text-align: left; font-weight: 600; }');
    SB.AppendLine('  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }');
    SB.AppendLine('  td.ctr, th.ctr { text-align: center; }');
    SB.AppendLine('  tr.sum td { font-weight: 700; background: #fff4d6; }');
    SB.AppendLine('  .empty { color: #999; font-style: italic; padding: 20px; text-align: center; }');
    SB.AppendLine('  .firma { margin-bottom: 14px; line-height: 1.4; }');
    SB.AppendLine('  .firma-nazwa { font-weight: 700; font-size: 14px; }');
    SB.AppendLine('  @media print { body { margin: 8mm; } .no-print { display: none; } }');
    SB.AppendLine('</style>');
    SB.AppendLine('</head>');
    SB.AppendLine('<body>');

    SB.Append(CompanyHeaderHTML);
    SB.AppendLine('<h1>Raport sprzedaży (FV / KFV)</h1>');
    SB.AppendLine('<div class="meta">Zakres dat wystawienia: <b>' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateFrom)) + '</b> &ndash; <b>' +
      HtmlEscape(FormatDateTime('yyyy-mm-dd', ADateTo)) + '</b><br>' +
      'Wygenerowano: ' + HtmlEscape(FormatDateTime('yyyy-mm-dd hh:nn', Now)) + '</div>');

    if qDok.Eof then
    begin
      SB.AppendLine('<div class="empty">Brak dokumentów sprzedaży w wybranym zakresie dat.</div>');
    end
    else
    begin
      SB.AppendLine('<table>');
      SB.AppendLine('<tr>');
      SB.AppendLine('  <th class="ctr" style="width:40px">L.p.</th>');
      SB.AppendLine('  <th>Kontrahent</th>');
      SB.AppendLine('  <th>Numer faktury</th>');
      SB.AppendLine('  <th class="ctr">Z dnia</th>');
      SB.AppendLine('  <th>Forma płatności</th>');
      SB.AppendLine('  <th class="num">Kwota Netto</th>');
      SB.AppendLine('  <th class="num">Kwota VAT</th>');
      SB.AppendLine('  <th class="num">Kwota Brutto</th>');
      SB.AppendLine('</tr>');

      while not qDok.Eof do
      begin
        Inc(Lp);

        NrDok := VarToStrSafe(qDok.FieldByName('d_nazwa').Value);
        if NrDok = '' then
          NrDok := '(bez numeru, id=' + IntToStr(qDok.FieldByName('d_id').AsInteger) + ')';

        Kontrahent := VarToStrSafe(qDok.FieldByName('k_nazwa').Value);

        PlatnoscOpis := Trim(VarToStrSafe(qDok.FieldByName('d_platnosc').Value));
        if PlatnoscOpis = '' then
          PlatnoscOpis := '(nie określono)';

        SB.AppendLine('<tr>');
        SB.AppendLine('  <td class="ctr">' + IntToStr(Lp) + '</td>');
        SB.AppendLine('  <td>' + HtmlEscape(Kontrahent) + '</td>');
        SB.AppendLine('  <td>' + HtmlEscape(NrDok) + '</td>');
        if not qDok.FieldByName('d_wyst').IsNull then
          SB.AppendLine('  <td class="ctr">' +
            FormatDateTime('yyyy-mm-dd', qDok.FieldByName('d_wyst').AsDateTime) + '</td>')
        else
          SB.AppendLine('  <td class="ctr"></td>');
        SB.AppendLine('  <td>' + HtmlEscape(PlatnoscOpis) + '</td>');
        SB.AppendLine('  <td class="num">' + FmtMoney(qDok.FieldByName('d_netto').Value) + '</td>');
        SB.AppendLine('  <td class="num">' + FmtMoney(qDok.FieldByName('d_vat').Value) + '</td>');
        SB.AppendLine('  <td class="num">' + FmtMoney(qDok.FieldByName('d_brutto').Value) + '</td>');
        SB.AppendLine('</tr>');

        if not qDok.FieldByName('d_netto').IsNull then
          TotalNetto := TotalNetto + qDok.FieldByName('d_netto').AsFloat;
        if not qDok.FieldByName('d_vat').IsNull then
          TotalVat := TotalVat + qDok.FieldByName('d_vat').AsFloat;
        if not qDok.FieldByName('d_brutto').IsNull then
          TotalBrutto := TotalBrutto + qDok.FieldByName('d_brutto').AsFloat;

        qDok.Next;
      end;

      // Wiersz podsumowania.
      SB.AppendLine('<tr class="sum">');
      SB.AppendLine('  <td colspan="5" style="text-align:right">RAZEM (dokumentów: ' +
        IntToStr(Lp) + '):</td>');
      SB.AppendLine('  <td class="num">' + FmtMoney(TotalNetto) + '</td>');
      SB.AppendLine('  <td class="num">' + FmtMoney(TotalVat) + '</td>');
      SB.AppendLine('  <td class="num">' + FmtMoney(TotalBrutto) + '</td>');
      SB.AppendLine('</tr>');
      SB.AppendLine('</table>');
      SB.AppendLine('<div class="meta">Uwaga: zestawienie obejmuje faktury sprzedaży (FV) oraz ich ' +
        'korekty (KFV). Proformy (FPV) nie są dokumentami sprzedaży i nie są uwzględniane.</div>');
    end;

    SB.AppendLine('</body>');
    SB.AppendLine('</html>');

    Result := SB.ToString;
  finally
    qDok.Free;
    SB.Free;
  end;
end;

function TfrmMain.BuildReportHTML(ADateFrom, ADateTo: TDateTime): string;
begin
  // cmbReportType.ItemIndex: 0 = zakupy (FZ/KFZ), 1 = sprzedaż (FV/KFV).
  if cmbReportType.ItemIndex = 1 then
    Result := BuildSalesReportHTML(ADateFrom, ADateTo)
  else
    Result := BuildPurchaseReportHTML(ADateFrom, ADateTo);
end;

function TfrmMain.ReportFileTag: string;
begin
  // Używany do nazw plików/tymczasowych - odróżnia zakupy od sprzedaży.
  if cmbReportType.ItemIndex = 1 then
    Result := 'sprzedazy'
  else
    Result := 'zakupow';
end;

function TfrmMain.CompanyHeaderHTML: string;
var
  AdresLine: string;
  SB: TStringBuilder;
begin
  Result := '';
  if not FFirmaZaladowana then
    Exit;
  if (FFirmaNazwa = '') and (FFirmaUlica = '') and
     (FFirmaMiejscowosc = '') and (FFirmaNIP = '') then
    Exit;

  // Adres sklejony jak w istniejącym raporcie zakupów: "ulica nrdomu/nrlokalu, kod miasto".
  AdresLine := Trim(FFirmaUlica + ' ' + FFirmaNrDomu);
  if FFirmaNrLokalu <> '' then
    AdresLine := AdresLine + '/' + FFirmaNrLokalu;
  if (FFirmaKodPoczt <> '') or (FFirmaMiejscowosc <> '') then
  begin
    if AdresLine <> '' then
      AdresLine := AdresLine + ', ';
    AdresLine := AdresLine + Trim(FFirmaKodPoczt + ' ' + FFirmaMiejscowosc);
  end;

  SB := TStringBuilder.Create;
  try
    SB.AppendLine('<div class="firma">');
    if FFirmaNazwa <> '' then
      SB.AppendLine('  <div class="firma-nazwa">' + HtmlEscape(FFirmaNazwa) + '</div>');
    if AdresLine <> '' then
      SB.AppendLine('  <div>' + HtmlEscape(AdresLine) + '</div>');
    if FFirmaNIP <> '' then
      SB.AppendLine('  <div>NIP: ' + HtmlEscape(FFirmaNIP) + '</div>');
    SB.AppendLine('</div>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

procedure TfrmMain.btnGenerateClick(Sender: TObject);
var
  D1, D2: TDateTime;
begin
  if Trim(edtFile.Text) = '' then
  begin
    ShowMessage('Wskaż najpierw plik bazy danych.');
    Exit;
  end;
  D1 := DateOf(dtpFrom.Date);
  D2 := DateOf(dtpTo.Date);
  if D1 > D2 then
  begin
    ShowMessage('Data "od" jest późniejsza niż data "do".');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try
    SetStatus('Łączenie z bazą...');
    ConnectToDatabase(edtFile.Text);

    SetStatus('Generowanie raportu...');
    FLastHTML := BuildReportHTML(D1, D2);

    memHTML.Lines.BeginUpdate;
    try
      memHTML.Text := FLastHTML;
    finally
      memHTML.Lines.EndUpdate;
    end;

    SetStatus(Format('Gotowe. Raport ma %d znaków. Możesz zapisać go do pliku lub otworzyć w przeglądarce.',
      [Length(FLastHTML)]));

    // Po pierwszym udanym wygenerowaniu raportu z danej ścieżki - zapisz ją w INI.
    // Dzięki temu działa też scenariusz "wkleiłem ścieżkę ręcznie" (bez btnBrowse).
    SaveSettings;
  except
    on E: Exception do
    begin
      SetStatus('Błąd: ' + E.Message);
      MessageDlg('Błąd podczas generowania raportu:'#13#10 + E.Message, mtError, [mbOK], 0);
    end;
  end;
  Screen.Cursor := crDefault;
end;

procedure TfrmMain.btnSaveClick(Sender: TObject);
begin
  if FLastHTML = '' then
  begin
    ShowMessage('Najpierw wygeneruj raport.');
    Exit;
  end;
  dlgSave.FileName := Format('raport_%s_%s_%s.html',
    [ReportFileTag,
     FormatDateTime('yyyymmdd', dtpFrom.Date),
     FormatDateTime('yyyymmdd', dtpTo.Date)]);
  if dlgSave.Execute then
  begin
    TFile.WriteAllText(dlgSave.FileName, FLastHTML, TEncoding.UTF8);
    // Zapamiętaj folder, w którym użytkownik zapisał raport - kolejny zapis
    // domyślnie powędruje tam samo.
    dlgSave.InitialDir := ExtractFilePath(dlgSave.FileName);
    SaveSettings;
    SetStatus('Zapisano: ' + dlgSave.FileName);
  end;
end;

procedure TfrmMain.btnOpenBrowserClick(Sender: TObject);
var
  TmpName: string;
begin
  if FLastHTML = '' then
  begin
    ShowMessage('Najpierw wygeneruj raport.');
    Exit;
  end;
  TmpName := TPath.Combine(TPath.GetTempPath,
    Format('raport_%s_%s.html', [ReportFileTag, FormatDateTime('yyyymmddhhnnss', Now)]));
  TFile.WriteAllText(TmpName, FLastHTML, TEncoding.UTF8);
  ShellExecute(Handle, 'open', PChar(TmpName), nil, nil, SW_SHOWNORMAL);
end;

{ Zewnętrzna funkcja WinAPI do (tymczasowej) zmiany domyślnej drukarki
  systemowej. Silnik przeglądarki przy "cichym" drukowaniu (bez okna
  "Drukuj") zawsze wysyła zadanie na aktualną drukarkę domyślną systemu -
  dlatego przełączamy ją na czas wydruku na "Microsoft Print to PDF",
  a zaraz potem przywracamy oryginalną. }
function SetDefaultPrinterW(pszPrinter: PWideChar): BOOL; stdcall;
  external 'winspool.drv' name 'SetDefaultPrinterW';

type
  // Prosta pomoc do bezpiecznego backupu/przywracania pojedynczej wartości
  // rejestru - pamięta też, czy wartość w ogóle istniała wcześniej (żeby po
  // przywróceniu nie zostawić czegoś, czego wcześniej nie było).
  TRegBackupValue = record
    Existed: Boolean;
    Value: string;
  end;

function BackupRegString(Reg: TRegistry; const AValueName: string): TRegBackupValue;
begin
  Result.Existed := Reg.ValueExists(AValueName);
  if Result.Existed then
    Result.Value := Reg.ReadString(AValueName)
  else
    Result.Value := '';
end;

procedure RestoreRegString(Reg: TRegistry; const AValueName: string; const ABackup: TRegBackupValue);
begin
  if ABackup.Existed then
    Reg.WriteString(AValueName, ABackup.Value)
  else if Reg.ValueExists(AValueName) then
    Reg.DeleteValue(AValueName);
end;

procedure TfrmMain.btnSavePdfClick(Sender: TObject);
const
  DM_PAPERSIZE = $0002;
  DMPAPER_A4   = 9;
  PDF_PRINTER_NAME = 'Microsoft Print to PDF';
  MAX_WAIT_MS = 10000;
  // Klucz rejestru używany przez silnik Trident/MSHTML (a więc i przez
  // TWebBrowser) do ustawień marginesów oraz nagłówka/stopki wydruku -
  // to te same ustawienia, które w starym Internet Explorerze były
  // dostępne w oknie "Ustawienia strony".
  IE_PAGESETUP_KEY = 'Software\Microsoft\Internet Explorer\PageSetup';
var
  TmpHtml: string;
  PdfIdx: Integer;
  OldPrinterName: string;
  Device, Driver, Port: string;
  hDMode: THandle;
  DevModeData: PDeviceMode;
  StartTick: Cardinal;
  vIn, vOut: OleVariant;
  Reg: TRegistry;
  HasPageSetupKey: Boolean;
  BakMarginL, BakMarginT, BakMarginR, BakMarginB, BakHeader, BakFooter: TRegBackupValue;
begin
  if FLastHTML = '' then
  begin
    ShowMessage('Najpierw wygeneruj raport.');
    Exit;
  end;

  PdfIdx := Printer.Printers.IndexOf(PDF_PRINTER_NAME);
  if PdfIdx < 0 then
  begin
    MessageDlg(
      'Nie znaleziono drukarki wirtualnej "' + PDF_PRINTER_NAME + '".'#13#10#13#10 +
      'W Windows 10/11 jest ona zwykle wbudowana - włącz ją w:'#13#10 +
      'Ustawienia -> Drukarki i skanery -> Dodaj urządzenie'#13#10 +
      '(albo: Panel sterowania -> Programy i funkcje -> Włącz/wyłącz funkcje ' +
      'systemu Windows -> "Microsoft Print to PDF").',
      mtError, [mbOK], 0);
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  SetStatus('Przygotowywanie pliku PDF...');
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    HasPageSetupKey := Reg.OpenKey(IE_PAGESETUP_KEY, True);
    if HasPageSetupKey then
    begin
      // Zapamiętaj obecne ustawienia, żeby przywrócić je dokładnie takimi,
      // jakie były (albo usunąć, jeśli wcześniej ich nie było).
      BakMarginL := BackupRegString(Reg, 'margin_left');
      BakMarginT := BackupRegString(Reg, 'margin_top');
      BakMarginR := BackupRegString(Reg, 'margin_right');
      BakMarginB := BackupRegString(Reg, 'margin_bottom');
      BakHeader  := BackupRegString(Reg, 'header');
      BakFooter  := BackupRegString(Reg, 'footer');

      // Zmniejsz marginesy (domyślnie 0.75") i wyczyść nagłówek/stopkę
      // (domyślnie tytuł+data u góry, adres URL na dole) - inaczej te
      // marginesy razem z szeroką tabelą raportu łatwo powodują ucięcie
      // prawej krawędzi na wydruku/PDF.
      Reg.WriteString('margin_left', '0.4');
      Reg.WriteString('margin_top', '0.4');
      Reg.WriteString('margin_right', '0.4');
      Reg.WriteString('margin_bottom', '0.4');
      Reg.WriteString('header', '');
      Reg.WriteString('footer', '');
    end;

    try
      // Zapisz raport do pliku tymczasowego - tak samo jak przy "Otwórz w przeglądarce".
      TmpHtml := TPath.Combine(TPath.GetTempPath,
        Format('raport_%s_%s.html', [ReportFileTag, FormatDateTime('yyyymmddhhnnss', Now)]));
      TFile.WriteAllText(TmpHtml, FLastHTML, TEncoding.UTF8);

      // Zapamiętaj bieżącą drukarkę domyślną, żeby przywrócić ją po wydruku.
      OldPrinterName := '';
      if Printer.Printers.Count > 0 then
        OldPrinterName := Printer.Printers[Printer.PrinterIndex];

      if not SetDefaultPrinterW(PWideChar(WideString(PDF_PRINTER_NAME))) then
        raise Exception.Create('Nie udało się ustawić "' + PDF_PRINTER_NAME +
          '" jako drukarki domyślnej.');
      Printer.PrinterIndex := PdfIdx;

      try
        // Wymuś format A4 dla tego wydruku - niezależnie od domyślnego ustawienia
        // sterownika/systemu (które bywa np. Letter na anglojęzycznych Windows).
        Printer.GetPrinter(Device, Driver, Port, hDMode);
        if hDMode <> 0 then
        begin
          DevModeData := GlobalLock(hDMode);
          try
            if Assigned(DevModeData) then
            begin
              DevModeData^.dmFields := DevModeData^.dmFields or DM_PAPERSIZE;
              DevModeData^.dmPaperSize := DMPAPER_A4;
            end;
          finally
            GlobalUnlock(hDMode);
          end;
          Printer.SetPrinter(Device, Driver, Port, hDMode);
        end;

        // Załaduj raport w ukrytej przeglądarce i poczekaj na pełne wczytanie
        // (drukowanie przed zakończeniem ładowania dałoby pusty/urwany wydruk).
        FWebBrowser.Navigate(TmpHtml);
        StartTick := GetTickCount;
        while (FWebBrowser.ReadyState <> READYSTATE_COMPLETE) and
              (GetTickCount - StartTick < MAX_WAIT_MS) do
          Application.ProcessMessages;

        // Drukuj "cicho" (bez okna dialogowego "Drukuj" przeglądarki) - sterownik
        // "Microsoft Print to PDF" i tak pokaże własne systemowe okno "Zapisz
        // wynik drukowania jako", w którym wskazuje się docelowy plik .pdf.
        SetStatus('Wskaż plik docelowy w oknie "Zapisz wynik drukowania jako"...');
        vIn := Unassigned;
        vOut := Unassigned;
        FWebBrowser.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, vIn, vOut);

        // WAŻNE: ExecWB wraca ze sterowaniem niemal natychmiast - samo
        // wysłanie zadania do spoolera (i wybranie AKTUALNEJ drukarki
        // domyślnej) dzieje się chwilę później, asynchronicznie. Jeśli
        // przywrócimy poprzednią drukarkę domyślną zbyt szybko, zadanie
        // "ucieknie" na starą drukarkę zamiast na "Microsoft Print to PDF" -
        // to najczęstsza przyczyna, gdy drukowanie mimo wszystko trafia na
        // drukarkę domyślną sprzed uruchomienia funkcji. Dajemy więc
        // spoolerowi margines czasu, zanim cokolwiek przywrócimy.
        StartTick := GetTickCount;
        while GetTickCount - StartTick < 5000 do
          Application.ProcessMessages;
      finally
        // Przywróć poprzednią drukarkę domyślną, niezależnie od wyniku.
        if OldPrinterName <> '' then
          SetDefaultPrinterW(PWideChar(WideString(OldPrinterName)));
      end;

      SetStatus('Gotowe. Zapisz plik w oknie "Zapisz wynik drukowania jako", jeśli się pojawiło.');
    except
      on E: Exception do
      begin
        SetStatus('Błąd: ' + E.Message);
        MessageDlg('Błąd podczas zapisywania do PDF:'#13#10 + E.Message, mtError, [mbOK], 0);
      end;
    end;
  finally
    // Przywróć oryginalne ustawienia marginesów/nagłówka/stopki, niezależnie
    // od tego, czy wydruk się powiódł.
    if HasPageSetupKey then
    begin
      RestoreRegString(Reg, 'margin_left', BakMarginL);
      RestoreRegString(Reg, 'margin_top', BakMarginT);
      RestoreRegString(Reg, 'margin_right', BakMarginR);
      RestoreRegString(Reg, 'margin_bottom', BakMarginB);
      RestoreRegString(Reg, 'header', BakHeader);
      RestoreRegString(Reg, 'footer', BakFooter);
      Reg.CloseKey;
    end;
    Reg.Free;
  end;
  Screen.Cursor := crDefault;
end;

end.

