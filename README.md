# Raport zakupów i sprzedaży - Delphi + FireDAC

Aplikacja desktopowa Delphi (VCL), która łączy się z bazą Microsoft Access (.mdb)
programu fakturującego i generuje czytelny raport HTML za podany zakres dat.
Do wyboru są dwa typy raportu (pole **Typ raportu** w górnym panelu):

- **Zakupy (FZ / KFZ)** - szczegółowy raport z pozycjami każdego dokumentu
  (nazwa towaru/usługi, ilość, cena netto/brutto) i podsumowaniem zbiorczym.
- **Sprzedaż (FV / KFV)** - zwięzłe zestawienie tabelaryczne, jeden wiersz na
  fakturę, z kolumnami: L.p., Kontrahent, Numer faktury, Z dnia, Forma
  płatności, Kwota Netto, Kwota VAT, Kwota Brutto - oraz wierszem
  podsumowania na końcu.

## Pliki projektu

- `RaportZakupow.dpr` - plik projektu
- `MainForm.pas` / `MainForm.dfm` - formularz główny
- `RaportZakupow.rc` - skrypt zasobów Windows (definiuje ikonę jako `MAINICON`)

## Wymagania

- Delphi 10.x lub nowsze (kod używa `var` w bloku, `TStringBuilder`, `System.IOUtils`).
  Zostało to przetestowane konceptualnie pod kątem Delphi 10.3+ / 11.x / 12.x.
- FireDAC (jest w komplecie Delphi Professional+).
- **Microsoft Access Database Engine** w wersji zgodnej z bitnością aplikacji:
  - aplikacja 32-bit -> Access ACE 32-bit, lub natywny sterownik Jet (Windows ma go z reguły fabrycznie dla .mdb 32-bit),
  - aplikacja 64-bit -> Access ACE 64-bit (do pobrania ze strony Microsoftu).
  Bez tego FireDAC zwróci błąd "Cannot find driver MSAcc".

## Jak otworzyć w Delphi

1. W IDE: *File -> Open Project...* -> wskaż `RaportZakupow.dpr`.
2. Skompiluj i uruchom (F9).

## Jak używać

1. Wskaż plik `.mdb` przyciskiem **Przeglądaj...** (jeśli plik `fakturka.mdb` leży obok exe,
   ścieżka uzupełni się automatycznie).
2. Ustaw zakres dat ręcznie albo skorzystaj z menu **Szybki wybór ▼**, w którym są:
   - Dzień dzisiejszy / wczorajszy
   - Bieżący / poprzedni miesiąc
   - Bieżący / poprzedni kwartał
   - Bieżący / poprzedni rok
   - Ostatnie 30 / 90 / 120 dni
3. Wybierz **Typ raportu**: Zakupy (FZ / KFZ) albo Sprzedaż (FV / KFV).
   Program zapamiętuje ostatnio wybrany typ w pliku INI.
4. Klik **Generuj raport** - wynik wyświetli się w polu tekstowym jako kod HTML.
5. **Zapisz HTML...** - zapisuje do pliku (nazwa domyślna zawiera typ raportu,
   np. `raport_sprzedazy_20260601_20260630.html`), **Otwórz w przeglądarce** -
   otwiera w domyślnej przeglądarce systemowej.
6. **Zapisz jako PDF (A4)** - drukuje wygenerowany raport bezpośrednio do PDF
   (patrz sekcja poniżej).

## Zapisz jako PDF (A4)

Przycisk konwertuje ostatnio wygenerowany raport HTML na PDF, bez żadnych
dodatkowych bibliotek - wykorzystuje wbudowaną w Windows wirtualną drukarkę
**"Microsoft Print to PDF"** oraz silnik przeglądarki (`TWebBrowser`) do
"cichego" wydruku (bez okna dialogowego "Drukuj").

Jak to działa:

1. Raport zapisywany jest do pliku tymczasowego HTML (jak przy "Otwórz w
   przeglądarce"), wczytywany do ukrytej przeglądarki wbudowanej w program.
2. Program tymczasowo przełącza domyślną drukarkę systemową na "Microsoft
   Print to PDF" (przeglądarka przy cichym drukowaniu zawsze wysyła zadanie
   na aktualną drukarkę domyślną) i wymusza dla niej **format A4**
   (niezależnie od domyślnego ustawienia sterownika, które bywa np. Letter
   na anglojęzycznych Windows).
3. Drukuje raport "po cichu" (bez okna "Drukuj" przeglądarki) - **pojawi się
   jednak natywne okno Windows "Zapisz wynik drukowania jako"**, w którym
   trzeba wskazać docelową nazwę i lokalizację pliku `.pdf`. Tego okna nie da
   się pominąć - to integralna część sterownika "Microsoft Print to PDF".
4. Po wydruku program przywraca poprzednią drukarkę domyślną.

Wymagania:

- Drukarka **"Microsoft Print to PDF"** musi być zainstalowana/włączona w
  Windows (jest wbudowana w Windows 10/11, zwykle domyślnie włączona). Jeśli
  jej nie ma, program pokaże komunikat z instrukcją, jak ją włączyć
  (Ustawienia -> Drukarki i skanery, albo Panel sterowania -> Włącz/wyłącz
  funkcje systemu Windows).
- Funkcja korzysta z `TWebBrowser` (silnik oparty o komponent Internet
  Explorer/Trident, dostępny standardowo w Delphi) - nie wymaga instalowania
  żadnych dodatkowych pakietów.

Jeśli mimo wszystko wydruk trafia na starą drukarkę domyślną zamiast na
"Microsoft Print to PDF":

- Program celowo odczekuje kilka sekund po wysłaniu zadania drukowania,
  zanim przywróci poprzednią drukarkę domyślną - `ExecWB` wraca ze
  sterowaniem niemal natychmiast, a samo wysłanie zadania do spoolera
  (z wybraniem *aktualnej* drukarki domyślnej) dzieje się chwilę później,
  asynchronicznie. Zbyt wczesne przywrócenie starej drukarki było
  pierwotną przyczyną tego problemu - jeśli nadal występuje na wolniejszym
  komputerze, można wydłużyć czas oczekiwania w `btnSavePdfClick`
  (stała `5000` - w milisekundach).
- Sprawdź w Windows: *Ustawienia -> Drukarki i skanery* - jeśli włączona
  jest opcja *"Zarządzaj moją drukarką domyślną"* ("Let Windows manage my
  default printer"), warto ją tymczasowo wyłączyć - przy tej opcji Windows
  może samodzielnie przełączać drukarkę domyślną w tle, niezależnie od
  tego, co ustawi program.

### Ucięta prawa krawędź tabeli na wydruku/PDF

Główną przyczyną okazał się rozmiar **ukrytej przeglądarki** używanej do
druku: silnik Trident/MSHTML wylicza CSS-owe szerokości procentowe (nasza
tabela ma `width: 100%`) względem rozmiaru samej kontrolki `TWebBrowser` w
pikselach - nie względem szerokości strony drukarki, nawet przy druku.
Kontrolka była celowo mała i niewidoczna (100 x 100 px), więc cała tabela
była faktycznie wyliczana (a potem drukowana) jako 100 pikseli szeroka -
większość kolumn i tak wychodziła poza ten wąski obszar i była ucinana.

Poprawka: kontrolka ma teraz rozmiar zbliżony do rzeczywistego obszaru
drukowalnego strony A4 (750 x 1060 px, czyli w przybliżeniu szerokość A4
pomniejszona o zmniejszone marginesy, przy typowych 96 dpi) - nadal
niewidoczna i poza obszarem roboczym formularza, ale z realistycznym
rozmiarem do celów liczenia układu tabeli.

Dodatkowo, dla większego marginesu bezpieczeństwa, program:

- Wymusza w CSS każdej tabeli raportu `table-layout: fixed` oraz
  zawijanie długiego tekstu (`word-wrap`/`overflow-wrap: break-word`) w
  komórkach - dzięki temu długa nazwa kontrahenta czy towaru zawija się do
  kolejnej linii zamiast rozciągać tabelę szerzej niż strona.
- Przed wydrukiem tymczasowo zmniejsza marginesy strony (domyślnie 0.75"
  z każdej strony) do 0.4" oraz czyści domyślny nagłówek/stopkę (tytuł,
  data, adres URL) w ustawieniach silnika druku (rejestr:
  `HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup`) -
  zyskuje to dodatkową szerokość na wydruku. Oryginalne ustawienia są
  zapamiętywane i przywracane zaraz po wydruku.

Jeśli mimo to któraś kolumna nadal się ucina (np. przy bardzo długich
nazwach towarów w raporcie zakupów bez spacji, których nie da się zawinąć),
można spróbować zwiększyć `FWebBrowser.Width`/`Height` w `FormCreate` (np.
do 900/1270) albo zmniejszyć rozmiar czcionki w CSS raportu
(`body { font-size: ... }`).

## Co znajdzie się w raporcie

### Raport zakupów (FZ / KFZ)

Dla każdej faktury zakupu z wybranego zakresu dat wystawienia:

- **Nagłówek dokumentu**: skrót typu (FZ/KFZ), pełny numer, data wystawienia,
  nazwa kontrahenta, NIP, adres.
- **Tabela pozycji**: lp., nazwa towaru/usługi, ilość, j.m., cena netto, cena brutto.
- **Podsumowanie dokumentu**: netto / VAT / brutto, z informacją o walucie.

Na końcu - **podsumowanie zbiorcze**: łączna liczba dokumentów i pozycji
oraz suma netto/VAT/brutto. Suma zbiorcza obejmuje wyłącznie dokumenty w PLN
(mieszanie walut nie miałoby sensu - dokumenty walutowe są widoczne na liście,
ale nie są sumowane do "razem").

### Raport sprzedaży (FV / KFV)

Jedna tabela, wiersz na każdą fakturę sprzedaży (bez rozbicia na pozycje), z kolumnami:

| Kolumna          | Pochodzenie                     |
|------------------|----------------------------------|
| L.p.             | licznik wierszy                  |
| Kontrahent       | `dok.nazwak`                     |
| Numer faktury    | `dok.nazwa`                      |
| Z dnia           | `dok.wyst` (data wystawienia)    |
| Forma płatności  | `dok.platnosc`                   |
| Kwota Netto      | `dok.netto`                      |
| Kwota VAT        | `dok.vat`                        |
| Kwota Brutto     | `dok.brutto`                     |

Na końcu tabeli - wiersz podsumowania z liczbą dokumentów oraz sumą
netto/VAT/brutto.

Typ dokumentu jest tu rozpoznawany przez `JOIN typdok td ON td.id = d.typ`
i filtr `td.skrot IN ('FV', 'KFV')` - w odróżnieniu od raportu zakupów, gdzie
`typ = 20/21` zostały ustalone wprost z analizy bazy. Rozpoznawanie po skrócie
jest bezpieczniejsze, bo nie zależy od konkretnych numerów `id` w `typdok`.

**Proformy (FPV) są celowo pominięte** - to dokumenty informacyjne, a nie
faktury sprzedaży w rozumieniu VAT. Jeśli mają się jednak znaleźć w raporcie,
wystarczy dopisać `'FPV'` do listy w klauzuli `IN (...)` w `BuildSalesReportHTML`.

Raport sprzedaży, w odróżnieniu od raportu zakupów, **nie rozbija sum według
waluty** - kwoty netto/VAT/brutto są sumowane wprost z bazy. Jeśli w danych
trafiają się faktury sprzedaży w innej walucie niż PLN, warto to doprecyzować.

### Nagłówek z danymi firmy

Oba raporty (zakupów i sprzedaży) pokazują na samej górze dane firmy z
tabeli `Firma`: `Nazwa`, `ulica`, `nrdomu`, `nrlokalu`, `kodpoczt`,
`miejscowosc`, `NIP`. Dane są wczytywane raz, automatycznie, zaraz po
połączeniu z bazą (funkcja `LoadFirmaData`) - nie ma potrzeby niczego
wpisywać ręcznie. Adres sklejany jest w formie "ulica nrdomu/nrlokalu,
kodpoczt miejscowość", tak samo jak adres kontrahenta w raporcie zakupów.

Jeśli tabela `Firma` nie istnieje w danym pliku bazy albo jest pusta, program
nie przerywa działania - nagłówek po prostu nie pokaże danych firmy.

## Plik ustawień (.ini)

Program zapamiętuje ostatnio używany plik bazy, foldery dialogów i wybrany
typ raportu w pliku `.ini`. Który plik jest używany, zależy od sposobu
uruchomienia:

- **Z parametrem** - `RaportZakupow.exe "C:\Sciezka\MojaFirma.ini"` - program
  używa dokładnie wskazanego pliku. Przydatne, gdy jednym programem obsługuje
  się kilka baz/firm, każda z osobną konfiguracją (osobny skrót na pulpicie
  z innym parametrem dla każdej).
- **Bez parametru** - program używa pliku `MFRaporty.ini` leżącego obok pliku
  `.exe`.

Aktualnie używana ścieżka do pliku INI jest pokazywana w pasku statusu zaraz
po uruchomieniu programu.

## Założenia oparte na strukturze bazy

- Faktury zakupu mają w tabeli `dok` pole `typ = 20` (FZ) lub `typ = 21` (KFZ);
  mapowanie jest sprawdzane przez join z tabelą `typdok`.
- Kwoty dokumentu pobierane są z pól `netto`, `vat`, `brutto` tabeli `dok`
  (pola `*_zakup` są w tej bazie wyzerowane dla FZ - program traktuje cenę
  z faktury zakupu jako podstawową cenę dokumentu).
- Pozycje dokumentu są w tabeli `dane`, klucz: `dane.nrdok = dok.id`.
- Pomijane są dokumenty oznaczone jako usunięte (`status = 1`).
- Filtr daty działa na polu `dok.wyst` (data wystawienia) - bezpiecznie,
  niezależnie od godziny w polu, dzięki użyciu `>= dataOd AND < (dataDo + 1 dzień)`.

## Drobne uwagi techniczne

- `TFDConnection` jest skonfigurowany w runtime (nie wymaga pre-definicji w FDExplorer).
- Driver: `MSAcc` (działa zarówno dla `.mdb` jak i `.accdb`, jeśli jest zainstalowany ACE).
- HTML jest zapisywany w UTF-8.
- Ukryta przeglądarka użyta do eksportu PDF (`TWebBrowser`, jednostka `SHDocVw`)
  jest tworzona w kodzie (`FormCreate`), a nie w pliku `.dfm` - to standardowy
  sposób unikania ręcznej edycji binarnych danych ActiveX w formularzu.
