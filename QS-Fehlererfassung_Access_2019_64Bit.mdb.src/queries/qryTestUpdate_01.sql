SELECT
  Mastertab.ID,
  Mastertab.LQEinheitID,
  ReklMemo.Lagerpüfung
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
