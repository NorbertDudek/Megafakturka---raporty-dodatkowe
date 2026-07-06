# Raport zakupów - Delphi + FireDAC

Aplikacja desktopowa Delphi (VCL), która łączy się z bazą Microsoft Access (.mdb)
programu fakturującego, wyciąga z niej **faktury zakupu (FZ) i korekty zakupu (KFZ)**
za podany zakres dat i generuje czytelny raport HTML.

## Pliki projektu

- `RaportZakupow.dpr` - plik projektu
- `MainForm.pas` / `MainForm.dfm` - formularz główny
- `RaportZakupow.ico` - ikona aplikacji (7 rozmiarów: 16, 24, 32, 48, 64, 128, 256 px)
- `RaportZakupow.rc` - skrypt zasobów Windows (definiuje ikonę jako `MAINICON`)

## Ikona aplikacji

Jak podpiąć ikonę pod plik exe - dwie opcje:

**Opcja A (najprostsza, w IDE Delphi):**
1. Otwórz projekt w Delphi.
2. *Project → Options → Application → Load Icon...* → wskaż `RaportZakupow.ico`.
3. *Save Project*. IDE zaktualizuje plik `RaportZakupow.res` w tle.

**Opcja B (z linii poleceń, bez klikania w IDE):**
1. Z folderu projektu uruchom: `brcc32.exe RaportZakupow.rc` (`brcc32` jest w `bin/` Delphi - jeśli nie ma w PATH, podaj pełną ścieżkę).
2. Powstanie `RaportZakupow.res`.
3. Skompiluj projekt (np. `dcc32 RaportZakupow.dpr`) - Delphi automatycznie zlinkuje `RaportZakupow.res` dzięki dyrektywie `{$R *.res}`.

Ikona w pliku `.ico` zawiera 7 rozmiarów (16, 24, 32, 48, 64, 128, 256 px), więc Windows zawsze pokaże najlepiej dopasowany wariant - od paska zadań po duże miniatury w eksploratorze.

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
3. Klik **Generuj raport** - wynik wyświetli się w polu tekstowym jako kod HTML.
4. **Zapisz HTML...** - zapisuje do pliku, **Otwórz w przeglądarce** - otwiera w domyślnej
   przeglądarce systemowej.

## Co znajdzie się w raporcie

Dla każdej faktury zakupu (FZ/KFZ) z wybranego zakresu dat wystawienia:

- **Nagłówek dokumentu**: skrót typu (FZ/KFZ), pełny numer, data wystawienia,
  nazwa kontrahenta, NIP, adres.
- **Tabela pozycji**: lp., nazwa towaru/usługi, ilość, j.m., cena netto, cena brutto.
- **Podsumowanie dokumentu**: netto / VAT / brutto, z informacją o walucie.

Na końcu - **podsumowanie zbiorcze**: łączna liczba dokumentów i pozycji
oraz suma netto/VAT/brutto. Suma zbiorcza obejmuje wyłącznie dokumenty w PLN
(mieszanie walut nie miałoby sensu - dokumenty walutowe są widoczne na liście,
ale nie są sumowane do "razem").

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
