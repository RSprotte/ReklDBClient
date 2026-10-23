TRANSFORM
  Count(
    Abfrage_Fehlergruppen_Fehlertexte.ID
  ) AS [Anzahl von ID]
SELECT
  Abfrage_Fehlergruppen_Fehlertexte.Fehlergruppe,
  Count(
    Abfrage_Fehlergruppen_Fehlertexte.ID
  ) AS Anzahl
FROM
  Abfrage_Fehlergruppen_Fehlertexte
GROUP BY
  Abfrage_Fehlergruppen_Fehlertexte.Fehlergruppe
PIVOT
  Format([Auftragsdatum], "mmm") In (
    "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun",
    "Jul", "Aug", "Sep", "Okt", "Nov",
    "Dez"
  );
