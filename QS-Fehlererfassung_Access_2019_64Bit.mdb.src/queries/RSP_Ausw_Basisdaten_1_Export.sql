SELECT
  RSP_Ausw_Basisdaten_1_1.*
FROM
  RSP_Ausw_Basisdaten_1_1
WHERE
  (
    (
      [Firma_Name] in ('INDEX-Werke GmbH & Co. KG')
    )
    And ([extern] = True)
    And (
      [VorgangstypID] in (7)
    )
    And (
      [Auftragsdatum] between #11/18/2025#
      and #11/18/2026#
    )
  );
