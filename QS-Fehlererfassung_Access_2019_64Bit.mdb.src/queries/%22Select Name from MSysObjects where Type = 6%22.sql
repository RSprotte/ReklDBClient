SELECT
  MSysObjects.Name
FROM
  MSysObjects
WHERE
  (
    (
      (MSysObjects.Type) = 6
    )
  )
GROUP BY
  MSysObjects.Name;
