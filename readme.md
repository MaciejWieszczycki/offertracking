# Raport Ofertowy – Dokumentacja / Offer Report – Documentation

## Polska Wersja

### Opis
Ten arkusz Excel automatyzuje codzienne raportowanie ofert oraz analizę przyczyn braku ofert. Makro VBA o nazwie `ZamknijDzienIRozlicz` przenosi dane z formularza (Sheet1) do bazy danych (Sheet2) i resetuje formularz, zachowując stałą strukturę 80 wierszy z numeracją. Dodatkowo, w Sheet1 znajduje się formuła, która na bieżąco oblicza % ofertowania.

### Struktura Arkusza
- **Sheet1 (Formularz):**
  - Tabela z 80 wierszami, zawierająca:
    - Kolumna 1: Numer połączenia (Lp.)
    - Kolumna 2: Checkbox "Była oferta?" (wartość TRUE/FALSE)
    - Kolumna 3: Powód braku oferty (tekst)
- **Sheet2 (Baza Danych):**
  - Tabela nazwana `BazaDanych` przechowująca dane:
    - Kolumna 1: Numer rozmowy
    - Kolumna 2: Status oferty (TRUE/FALSE)
    - Kolumna 3: Powód braku oferty
    - Kolumna 4: Data

### Jak Działa Makro
1. Makro iteruje przez wszystkie wiersze formularza w Sheet1.
2. Dla wierszy, gdzie pole „Powód braku” jest niepuste, dane są przenoszone do tabeli `BazaDanych` w Sheet2, a numer rozmowy jest ustalany automatycznie.
3. Makro zlicza liczbę wierszy, gdzie checkbox jest zaznaczony (TRUE), oraz gromadzi powody braku oferty (pomijając wpisy „ZAOFERTOWANO”).
4. Po przeniesieniu danych makro resetuje zawartość kolumn 2 i 3 (checkbox ustawiony na FALSE i pole „Powód” czyszczone) w formularzu, pozostawiając numerację (kolumna 1) bez zmian.
5. Wyświetlane jest podsumowanie dnia zawierające statystyki ofertowania dziennego i ogólnego oraz najczęstszy powód braku oferty.

### Formuła na % Ofertowania (Live Percentage Calculation)
W Sheet1 umieść poniższą formułę (zakresy dostosuj do swojej tabeli):

*Ustaw format komórki na procentowy.*

### Instrukcje Użytkowania
1. **Konfiguracja makra:**  
   - Otwórz edytor VBA (Alt + F11), wstaw nowy moduł i wklej kod makra `ZamknijDzienIRozlicz`.
2. **Przycisk:**  
   - Na Sheet1 dodaj przycisk (Form Control) z zakładki Developer i przypisz do niego makro.
3. **Wypełnianie danych:**  
   - Uzupełniaj formularz w Sheet1. Po zakończeniu dnia kliknij przycisk, aby przenieść dane do Sheet2 i zresetować formularz.
4. **Formuła:**  
   - Upewnij się, że komórka z formułą na % ofertowania jest poprawnie umieszczona i sformatowana.

### Wymagania
- Excel z włączonymi makrami.
- Struktura arkusza zgodna z powyższym opisem.
- Tabela w Sheet2 musi być nazwana **BazaDanych**.

---

## English Version

### Overview
This Excel workbook automates the daily reporting of offers and analyzes the reasons for no offer. The VBA macro `ZamknijDzienIRozlicz` transfers data from a form in Sheet1 to a database in Sheet2 and resets the form while maintaining a fixed structure of 80 rows with sequential numbering. Additionally, a live formula in Sheet1 calculates the current offer percentage.

### Worksheet Structure
- **Sheet1 (Form):**
  - A table with 80 rows containing:
    - Column 1: Connection number (Sequence)
    - Column 2: "Offer Given?" Checkbox (Boolean TRUE/FALSE)
    - Column 3: Reason for no offer (text)
- **Sheet2 (Database):**
  - A table named `BazaDanych` storing:
    - Column 1: Conversation number
    - Column 2: Offer status (TRUE/FALSE)
    - Column 3: Reason for no offer
    - Column 4: Date

### How the Macro Works
1. The macro loops through all rows in the form table on Sheet1.
2. For rows with a non-empty "Reason" field, the data is transferred to the `BazaDanych` table in Sheet2, and the conversation number is automatically assigned.
3. The macro counts the number of rows with the checkbox marked TRUE (offer given) and collects reasons for no offer (ignoring entries marked "ZAOFERTOWANO").
4. After transferring the data, the macro resets columns 2 and 3 (unchecks the checkbox and clears the reason) in the form, leaving the sequence numbers intact.
5. A summary message box displays daily and overall offer statistics, along with the most common reason for no offer.

### Live Offer Percentage Formula
Place the following formula in a cell on Sheet1 (adjust ranges as needed):

*Format the cell as a percentage.*

### User Instructions
1. **VBA Macro Setup:**  
   - Open the VBA editor (Alt + F11), insert a new module, and paste the `ZamknijDzienIRozlicz` macro code.
2. **Button Assignment:**  
   - On Sheet1, add a Form Control button and assign the macro to it.
3. **Data Entry:**  
   - Fill in the form on Sheet1. At the end of the day, click the button to transfer data to Sheet2 and reset the form.
4. **Formula:**  
   - Ensure the live offer percentage formula is placed in the appropriate cell.

### Requirements
- Excel with macros enabled.
- Worksheet structure as described above.
- The table in Sheet2 must be named **BazaDanych**.

