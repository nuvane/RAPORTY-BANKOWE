<h1>Generator Sald Dziennych</h1>

Projekt oparty na Pythonie i bibliotece openpyxl oraz matplotlib.<br>
Automatyczny generator raportu sald dziennych z kont bankowych dla wielu spółek i walut.<br> 
Skrypt tworzy raport Excel (<b>SaldoDzienny.xlsx</b> do pobrania) zawierający dane, wykresy trendów oraz arkusz podsumowujący.<br>

<hr>

<h2>Funkcje </h2>

- Obsługuje wiele firm i walut (np. Company A / Company B, PLN / EUR)
- Generowanie danych (symulowanych) od początku roku
- Eksport do Excela z:
  - Arkuszami dla każdej firmy i waluty
  - Wykresami trendów sald (Trendy sald od początku roku i dla ostatnich 3 miesięcy) dla każdego konta
  - Zbiorczym arkuszem podsumowującym z łącznymi saldami

<hr>

<h3>Arkusz: Company A – PLN</h3>
<img src="images/company%20A%20PLN.png" alt="Arkusz Company A" width="700"/>

<hr>

<h3>Trend sald: Company A – EUR</h3>
<img src="images/wykres%20A%20EUR.png" alt="Wykres EUR A" width="700"/>

<hr>

<h3>Podsumowanie sald: wszystkie spółki – PLN</h3>
<img src="images/podsumowanie%20PLN.png" alt="Podsumowanie PLN" width="700"/>