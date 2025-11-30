## Informacja prawna
Projekt powstał na podstawie materiałów ze szkolenia Udemy. 
Kod został zmodyfikowany i dostosowany przeze mnie, jednak pełne prawa 
autorskie do materiałów szkoleniowych należą do autora kursu - Kyle Pew.
Repozytorium służy jedynie jako prezentacja umiejętności.


## Automatyzacja raportów w Excelu z wykorzystaniem VBA - projekt szkoleniowy Udemy.com

Projekt: QuarterlyReport.xlsm


Projekt przedstawia kompletną automatyzację raportu kwartalnego w Excelu
z wykorzystaniem makr w języku **VBA**.  
Celem jest szybkie łączenie danych z wielu arkuszy, czyszczenie danych oraz
generowanie gotowego raportu rocznego (arkusz **YEARLY REPORT**) jednym kliknięciem.

Projekt doskonale odzwierciedla realne zadania analityczne wykonywane
w działach raportowania, finansów, controllingu oraz analityki sprzedaży.

---

## Struktura repozytorium

```plaintext excel-vba-automation/
├── README.md
├── LICENSE
├── QuarterlyReport.xlsm
├── src/
    └── macros.bas
├── data/
    └── sample.xlsm
├── plots/
    └── done.png

```

## Funkcjonalność projektu

Makra automatyzują:

### 1. Czyszczenie danych
- usuwanie pustych wierszy,
- usuwanie zbędnego formatowania,
- przygotowanie tabeli do agregacji.

### 2. Generowanie raportu rocznego
Makro **CreateYearlyReport**:
- przechodzi przez wszystkie arkusze w pliku (poza YEARLY REPORT),
- kopiuje dane z zakresu od `A2` do ostatniego niepustego wiersza,
- wkleja je kolejno do arkusza **YEARLY REPORT**,
- wstawia nagłówki (tylko raz),
- tworzy sumy w kolumnach (Total – kolumna F).

### 3. Automatyczne formatowanie
- nagłówki są pogrubione i mają kolor tła,
- kolumny C–F są formatowane jako waluty,
- sheet jest automatycznie dopasowany (`AutoFit`).

---

## Wykorzystane technologie

- **Excel – poziom zaawansowany**
- **VBA – makra do automatyzacji**
- Formatowanie tabel
- Funkcje Excela: SUM, SUMIF, formatowanie warunkowe
- Obsługa wielu arkuszy i dynamicznego wyszukiwania zakresów

---

## Zawartość repozytorium

- **QuarterlyReport.xlsm** – główny plik z makrami i przykładowymi danymi
- **README.md** – dokumentacja projektu

---

## Makra i procedury w projekcie

Projekt zawiera poniższe procedury:

- **CreateYearlyReport** – główne makro generujące raport roczny  
- **AutomateTotalSUM_OnSheet** – automatyczne obliczanie sumy w kolumnie `Total`  
- **InsertHeadersToSheet** – dodanie nagłówków do arkusza  
- **FormatHeadersOnSheet** – formatowanie nagłówków oraz kolumn (waluta, AutoFit)  

Wszystkie makra są refaktoryzowane i **nie używają `.Select` ani `ActiveCell`**, co czyni je stabilnymi i szybkim w działaniu.

---

## Jak uruchomić projekt (krok po kroku)

1. Pobierz plik **data/sample.xlsm** oraz src/macro.bas
2. Otwórz go sample.xlsm w Excelu.
3. Włącz obsługę makr:  
   **Plik → Opcje → Centrum zaufania → Ustawienia Centrum zaufania → Ustawienia makr → Włącz makra.**
4. Włącz zakładkę **Deweloper**:  
   Plik → Opcje → Dostosuj wstążkę → zaznacz „Deweloper”.
5. Importuj makro:  
   **Deweloper → Visual basic → Import File → wybierz `macro.bas`**
6. Uruchom makro:  
   **Deweloper → Makra → wybierz `CreateYearlyReport` → Uruchom**
7. Gotowe!  
   Wyniki pojawią się w arkuszu **YEARLY REPORT**.

---

## Przykładowy efekt działania makra

Po uruchomieniu:
- dane z wszystkich arkuszy są połączone,
- nagłówki dodane tylko raz,
- kolumny przeliczone i sformatowane,
- w kolumnie F znajduje się suma kwartalna (`SUM(F2:F...)`).

---

## Uwagi i dobre praktyki

- Plik zawiera wyłącznie dane **przykładowe** – nie zawiera danych wrażliwych.  
- Zakładamy, że dane w każdym arkuszu zaczynają się od komórki **A2**.  
- Arkusz `YEARLY REPORT` jest pomijany podczas scalania danych.  
- Makro samo utworzy arkusz `YEARLY REPORT`, jeśli jeszcze nie istnieje.  
- Logika makra została zoptymalizowana – działa szybciej niż standardowe nagrywane makra.

---

## Potencjalne problemy i rozwiązania

- **Makra się nie uruchamiają** – sprawdź czy włączone są makra (Centrum zaufania).
- **Błąd „Subscript out of range”** – upewnij się, że nie zmieniłeś nazwy arkuszy.
- **Puste wyniki** – sprawdź czy w arkuszach są dane w kolumnie A od wiersza 2.

---
