SELECT
  RSP_Ausw_Basisdaten_1_1.*
FROM
  RSP_Ausw_Basisdaten_1_1
WHERE
  (
    (
      [Kategorie-Status] in (
        'Bearbeitung: Service', 'zurück an: Service'
      )
    )
    And ([extern] = True)
    And (
      [Auftragsdatum] between #10/13/2025#
      and #10/13/2026#
    )
  );
