SELECT
  Mastertab.ID,
  Mastertab.[Material-Kosten],
  Mastertab.[Fracht-Kosten],
  Mastertab.Prüfkosten,
  Mastertab.[Belastungsbetrag Kunde],
  Mastertab.[Belastungsbetrag LQ],
  Mastertab.Reparaturbetrag,
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.Verursacher
FROM
  Mastertab;
