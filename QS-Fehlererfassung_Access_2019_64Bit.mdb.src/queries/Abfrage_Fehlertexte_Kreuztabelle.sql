TRANSFORM
  Count(
    Abfrage_Fehlergruppen_Fehlertexte.ID
  ) AS [Anzahl von ID]
SELECT
  Abfrage_Fehlergruppen_Fehlertexte.Fehlertexte,
  Count(
    Abfrage_Fehlergruppen_Fehlertexte.ID
  ) AS Anzahl
FROM
  Abfrage_Fehlergruppen_Fehlertexte
GROUP BY
  Abfrage_Fehlergruppen_Fehlertexte.Fehlertexte
PIVOT
  Format([Auftragsdatum], "mmm") In (
    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun",
    "Jul", "Aug", "Sep", "Okt", "Nov",
    "Dez"
  );
