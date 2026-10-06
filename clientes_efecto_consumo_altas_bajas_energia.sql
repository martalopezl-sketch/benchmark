CREATE OR REPLACE TABLE
  `mm-new-business-reporting.ENERGY.clientes_efecto_consumo_altas_bajas_energia` AS (
  WITH
    -- Código de tarifa -> nombre (las tablas de salida trabajan con el nombre)
    pricelist_names AS (
    SELECT
      code,
      MIN(name) AS name
    FROM
      `mm-datamart-kd.ENERGY.lucera_product_pricelist`
    GROUP BY
      code),
    -- Cada fila de sales_energy_mo es un tramo de tarifa de un contrato (sin solapes).
    -- Solo luz (E_SALE): se excluyen gas (G_*) y servicios (E_SVA / G_SVA).
    -- Alta = registered_activation_date, baja = registered_deactivation_date.
    tarifa_llamada AS (
    SELECT
      DISTINCT s.terminal_contract_id AS polissa_id,
      DATE(s.registered_activation_date) AS fecha_inicio_actual,
      DATE(s.registered_deactivation_date) AS fecha_anterior_final,
      CASE
        WHEN COALESCE(p.name, s.energy_pricelist_code) IN ('Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)', 'Fijo 24H Básica Retención Natural (Abril 2023 - 0,1288)') THEN 'Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)'
        ELSE COALESCE(p.name, s.energy_pricelist_code)
    END
      AS nombre_tarifa,
      s.energy_brand AS company_name,
      s.signup_id
    FROM
      `mm-datamart-kd.ENERGY.sales_energy_mo` s
    LEFT JOIN
      pricelist_names p
    ON
      p.code = s.energy_pricelist_code
    WHERE
      s.energy_pricelist_code LIKE 'E_SALE%'
      AND s.registered_activation_date IS NOT NULL
      AND DATE(s.registered_activation_date) <= DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY)),
    variables AS (
    SELECT
      CAST(DATE_TRUNC(CURRENT_DATE() - INTERVAL 5 YEAR, MONTH) AS date) AS start_date ),
    days AS (
    SELECT
      DATE_TRUNC(dd, day) AS day
    FROM
      UNNEST (GENERATE_DATE_ARRAY ((
          SELECT
            start_date
          FROM
            variables), current_date, INTERVAL 1 DAY)) dd ),
    months AS (
    SELECT
      DATE_TRUNC(dd,month) AS month
    FROM
      UNNEST (GENERATE_DATE_ARRAY ((
          SELECT
            start_date
          FROM
            variables), current_date, INTERVAL 1 MONTH)) dd ),
    month_days AS (
    SELECT
      m.month,
      days.day
    FROM
      days
    CROSS JOIN
      months AS m
    ORDER BY
      m.month,
      days.day ),
    activaciones AS (
    SELECT
      DATE_TRUNC(fecha_inicio_actual, month) AS activation_month,
      DATE_TRUNC(fecha_inicio_actual, day) AS activation_day,
      nombre_tarifa AS Tarifa_Alta,
      company_name AS tenant_code,
      COUNT(*) AS Contracts_Activated
    FROM
      tarifa_llamada
    GROUP BY
      1,
      2,
      3,
      4
    ORDER BY
      1,
      2,
      3),
    bajas_cierre AS (
    SELECT
      DATE_TRUNC(fecha_anterior_final, month) AS bajas_month,
      DATE_TRUNC(fecha_anterior_final, day) AS bajas_day,
      nombre_tarifa AS Tarifa_Baja,
      company_name AS tenant_code,
      count (*) AS Contracts_Churned
    FROM
      tarifa_llamada
    GROUP BY
      1,
      2,
      3,
      4
    ORDER BY
      1,
      2,
      3 ),
    bajas_table AS (
    SELECT
      DISTINCT *
    FROM (
      SELECT
        DISTINCT EXTRACT(year
        FROM
          month_days.month) AS Year,
        EXTRACT(month
        FROM
          month_days.day) AS Month,
        EXTRACT(day
        FROM
          month_days.day) AS Day,
        tenant_code,
        Tarifa_Baja,
        Contracts_Churned,
        EXTRACT(day
        FROM
          month_days.day) AS Dias_Activo_Bajas,
        Contracts_Churned*EXTRACT(day
        FROM
          month_days.day) AS Dias_Totales_Bajas,
        'Bajas' AS kpi_type
      FROM
        month_days
      LEFT JOIN
        bajas_cierre
      ON
        bajas_cierre.bajas_day = month_days.day
        AND bajas_cierre.bajas_month = month_days.month
      WHERE
        (Contracts_Churned IS NOT NULL)
      ORDER BY
        1,
        2,
        3 ) ),
    activaciones_table AS (
    SELECT
      DISTINCT *
    FROM (
      SELECT
        DISTINCT EXTRACT(year
        FROM
          month_days.month) AS Year,
        EXTRACT(month
        FROM
          month_days.day) AS Month,
        EXTRACT(day
        FROM
          month_days.day) AS Day,
        tenant_code,
        Tarifa_Alta,
        Contracts_Activated,
        EXTRACT(DAY
        FROM
          LAST_DAY(DATE(EXTRACT (YEAR
              FROM
                month_days.month), EXTRACT(MONTH
              FROM
                month_days.month), 1))) - EXTRACT(day
        FROM
          month_days.day) + 1 AS Dias_Activo_Altas,
        Contracts_Activated*(EXTRACT(day
          FROM
            month_days.day) - 1) AS Dias_Totales_Altas,
        'Altas' AS kpi_type
      FROM
        month_days
      LEFT JOIN
        activaciones
      ON
        activaciones.activation_day = month_days.day
        AND activaciones.activation_month = month_days.month
      WHERE
        (Contracts_Activated IS NOT NULL)
      ORDER BY
        1,
        2,
        3 ) ),
    dates_lab AS (
    SELECT
      DATE_TRUNC(CURRENT_DATE(), MONTH) AS start_date,
      DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY) AS end_date ),
    array_dates_lab AS (
    SELECT
      DATE_ADD(dates_lab.start_date, INTERVAL
      OFFSET
        DAY) AS date,
    FROM
      dates_lab
    CROSS JOIN
      UNNEST(GENERATE_ARRAY(0, DATE_DIFF(dates_lab.end_date, dates_lab.start_date, DAY))) AS
    OFFSET
      ),
    lab_days_table AS (
    SELECT
      COUNT(*) AS lab_days
    FROM (
      SELECT
        *,
        CASE
          WHEN laborable = 'Festivo' OR EXTRACT(DAYOFWEEK FROM date) BETWEEN 0 AND 2 OR laborable = 'Festivo' THEN 'No Laborable'
          ELSE 'Laborable'
      END
        AS day_type
      FROM
        array_dates_lab
      LEFT JOIN
        `mm-new-business-reporting.ENERGY.festivos` festivos
      ON
        date = CAST(festivos.fecha AS date) )
    WHERE
      day_type = 'Laborable'),
    promedio_festivos_bajas AS (
    SELECT
      year,
      month,
      day_type,
      tarifa_baja,
      tenant_code,
      kpi_type,
      AVG(CASE
          WHEN day_type = 'Laborable' THEN contracts_churned
          ELSE 0
      END
        ) AS avg_contracts_churned
    FROM (
      SELECT
        *,
        DATE(year, month, day) AS fecha_fixed,
        CASE
          WHEN laborable = 'Festivo' OR EXTRACT(DAYOFWEEK FROM DATE(year, month, day)) BETWEEN 0 AND 2 OR laborable = 'Festivo' THEN 'No Laborable'
          ELSE 'Laborable'
      END
        AS day_type
      FROM
        bajas_table
      LEFT JOIN
        `mm-new-business-reporting.ENERGY.festivos` festivos
      ON
        DATE(year, month, day) = CAST(festivos.fecha AS date))
    GROUP BY
      1,
      2,
      3,
      4,
      5,
      6 ),
    promedio_festivos_altas AS (
    SELECT
      year,
      month,
      day_type,
      tarifa_alta,
      tenant_code,
      kpi_type,
      SAFE_DIVIDE(SUM(CASE
          WHEN day_type = 'Laborable' THEN contracts_activated
          ELSE 0
      END
        ), lab_days) AS contracts_activated
    FROM (
      SELECT
        *,
        DATE(year, month, day) AS fecha_fixed,
        CASE
          WHEN laborable = 'Festivo' OR EXTRACT(DAYOFWEEK FROM DATE(year, month, day)) BETWEEN 0 AND 2 OR laborable = 'Festivo' THEN 'No Laborable'
          ELSE 'Laborable'
      END
        AS day_type
      FROM
        activaciones_table
      LEFT JOIN
        `mm-new-business-reporting.ENERGY.festivos` festivos
      ON
        DATE(year, month, day) = CAST(festivos.fecha AS date)
      CROSS JOIN
        lab_days_table)
    GROUP BY
      1,
      2,
      3,
      4,
      5,
      6,
      lab_days ),
    promedio_festivos AS (
    SELECT
      *
    FROM
      promedio_festivos_altas
    UNION ALL
    SELECT
      *
    FROM
      promedio_festivos_bajas),
    dates AS (
    SELECT
      CURRENT_DATE() AS start_date,
      LAST_DAY(CURRENT_DATE()) AS end_date ),
    array_dates AS (
    SELECT
      DATE_ADD(dates.start_date, INTERVAL
      OFFSET
        DAY) AS date,
    FROM
      dates
    CROSS JOIN
      UNNEST(GENERATE_ARRAY(0, DATE_DIFF(dates.end_date, dates.start_date, DAY))) AS
    OFFSET
      ),
    festivos_estimacion AS (
    SELECT
      *,
      CASE
        WHEN laborable = 'Festivo' OR EXTRACT(DAYOFWEEK FROM date) BETWEEN 0 AND 2 THEN 'No Laborable'
        ELSE 'Laborable'
    END
      AS day_type
    FROM
      array_dates
    LEFT JOIN
      `mm-new-business-reporting.ENERGY.festivos` festivos
    ON
      date = CAST(festivos.fecha AS date) ),
    kpis_table AS (
    SELECT
      *
    FROM
      activaciones_table
    UNION ALL
    SELECT
      *
    FROM
      bajas_table ),
    tarifas_a_promediar AS (
    SELECT
      DISTINCT year,
      month,
      tarifa_alta,
      kpi_type,
      tenant_code
    FROM
      kpis_table ),
    promedio_final AS (
    SELECT
      EXTRACT(YEAR
      FROM
        date) AS year,
      EXTRACT(MONTH
      FROM
        date) AS month,
      EXTRACT(DAY
      FROM
        date) AS day,
      tarifa_alta AS tarifa,
      tenant_code,
      kpi_type,
      contracts_activated
    FROM
      festivos_estimacion
    LEFT JOIN
      promedio_festivos
    ON
      festivos_estimacion.day_type = promedio_festivos.day_type
      AND EXTRACT(YEAR
      FROM
        date) = year
      AND EXTRACT(MONTH
      FROM
        date) = month ),
    tarifas_final AS (
    SELECT
      promedio_final.*
    FROM
      promedio_final
    LEFT JOIN
      tarifas_a_promediar
    ON
      promedio_final.tenant_code = tarifas_a_promediar.tenant_code
      AND promedio_final.tarifa = tarifas_a_promediar.tarifa_alta
      AND promedio_final.kpi_type = tarifas_a_promediar.kpi_type
      AND promedio_final.year = tarifas_a_promediar.year
      AND promedio_final.month = tarifas_a_promediar.month
    WHERE
      tarifas_a_promediar.tarifa_alta IS NOT NULL ),
    clientes_efecto_consumo_table AS (
    SELECT
      *,
      CASE
        WHEN kpi_type = 'Altas' THEN EXTRACT(DAY FROM LAST_DAY(DATE (year, month, day))) - DAY + 1
        WHEN kpi_type = 'Bajas' THEN day
        ELSE NULL
    END
      AS days_active,
      SAFE_DIVIDE((contracts *
        CASE
          WHEN kpi_type = 'Altas' THEN EXTRACT(DAY FROM DATE (year, month, day)) - 1
          WHEN kpi_type = 'Bajas' THEN day
          ELSE NULL
      END
        ), (EXTRACT(DAY
        FROM
          LAST_DAY(DATE(year, month, 1))))) AS clientes_efecto_consumo
    FROM (
      SELECT
        year,
        month,
        day,
        tarifa_alta AS tarifa,
        tenant_code,
        kpi_type,
        contracts_activated AS contracts
      FROM
        kpis_table
      UNION ALL
      SELECT
        year,
        month,
        day,
        tarifa,
        tenant_code,
        kpi_type,
        contracts_activated AS contracts
      FROM
        tarifas_final ))
  SELECT
    *
  FROM
    clientes_efecto_consumo_table )
