# Komnata

Komnata to prosta aplikacja webowa do zarządzania rzeczami znajdującymi się w sektorach i na regałach.

## Funkcje MVP
- dodawanie produktów z przypisaniem do sektora i regału,
- stałe listy wyboru sektorów i regałów,
- podgląd zawartości według lokalizacji,
- wyszukiwarka z podpowiedziami podobnych pozycji,
- wersja do wydruku,
- możliwość rozbudowy listy sektorów i regałów przez użytkownika.

## Domyślna konfiguracja
- sektory: `1 A` do `6 C`
- regały: `0A` do `5C`

## Uruchomienie lokalne
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Aplikacja uruchomi się pod adresem `http://127.0.0.1:5000`.

## Wdrożenie publiczne
### Railway
1. Załóż repozytorium GitHub i wrzuć pliki projektu.
2. Na Railway wybierz deploy z repozytorium GitHub.
3. Start command ustaw na:
```bash
gunicorn app:app
```
4. Wygeneruj publiczny domain w ustawieniach usługi.

### Render
Możesz wdrożyć jako Python Web Service. Dla produkcji lepiej użyć zewnętrznej bazy niż samego lokalnego pliku SQLite, bo trwałość lokalnych plików zależy od planu i konfiguracji. Render oferuje publiczne web services i integrację z repozytorium Git. citeturn708663search3turn708663search5

### Railway – dlaczego dobry na start
Railway ma oficjalny przewodnik dla Flask, obsługuje wdrożenie z GitHub, start command typu `gunicorn ...` i generowanie publicznej domeny z panelu. citeturn708663search1turn708663search13

## Kolejne ulepszenia, które warto dodać
- edycja istniejących pozycji,
- logowanie użytkowników i role,
- eksport do PDF/CSV,
- zdjęcia produktów,
- kody QR / kody kreskowe,
- historia zmian,
- osobne poziomy: podłoga / półka / pojemnik,
- API do integracji z innym systemem.
