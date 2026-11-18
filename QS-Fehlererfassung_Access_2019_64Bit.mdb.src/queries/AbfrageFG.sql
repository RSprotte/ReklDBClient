SELECT
  Mastertab.ID,
  Mastertab.LQWorkstationID,
  LQWorkstation.LQWorkstation
FROM
  Mastertab
  LEFT JOIN LQWorkstation ON Mastertab.LQWorkstationID = LQWorkstation.ID
WHERE
  (
    (
      (LQWorkstation.LQWorkstation) Like "FG 48 Platz*"
    )
    AND (
      (LQWorkstation.[ID-LQBereich]) = 54
    )
  )
ORDER BY
  Mastertab.ID DESC;
