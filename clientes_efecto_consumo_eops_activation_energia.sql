CREATE OR REPLACE TABLE `mm-new-business-reporting.ENERGY.clientes_efecto_consumo_eops_activation_energia` AS (
WITH clientes_efecto_consumo_table AS (
  SELECT * FROM `mm-new-business-reporting.ENERGY.clientes_efecto_consumo_altas_bajas_energia`
),

-- Código de tarifa -> nombre (el CEC y el Data Studio trabajan con el nombre)
pricelist_names AS (
  SELECT code, MIN(name) AS name
  FROM `mm-datamart-kd.ENERGY.lucera_product_pricelist`
  GROUP BY code
),

-- Cada fila de sales_energy_mo es un tramo de tarifa de un contrato (sin solapes).
-- Solo luz (E_SALE): se excluyen gas (G_*) y servicios (E_SVA / G_SVA).
-- Las filas change_type = 'change' son los cambios de tarifa del contrato.
tarifa_llamada AS (
  SELECT DISTINCT
    s.signup_id,
    s.terminal_contract_id AS polissa_id,
    s.energy_effective_dealer AS effective_dealer,
    s.energy_discount_code AS affiliate_code,
    DATE(s.registered_activation_date) AS data_alta,
    DATE(s.registered_deactivation_date) AS data_baixa,
    IF(s.change_type = 'change', DATE(s.registered_activation_date), NULL) AS fecha_ultimo_cambio_tarifa,
    CASE WHEN COALESCE(p.name, s.energy_pricelist_code) IN ('Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)', 'Fijo 24H Básica Retención Natural (Abril 2023 - 0,1288)')
         THEN 'Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)'
         ELSE COALESCE(p.name, s.energy_pricelist_code) END AS nombre_tarifa,
    s.energy_brand AS company_name
  FROM `mm-datamart-kd.ENERGY.sales_energy_mo` s
  LEFT JOIN pricelist_names p ON p.code = s.energy_pricelist_code
  WHERE s.energy_pricelist_code LIKE 'E_SALE%'
    AND s.registered_activation_date IS NOT NULL
    AND DATE(s.registered_activation_date) <= DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY)
),

base_calculo AS (
    SELECT
      FORMAT_DATE('%Y%m', fecha) AS yearmonth,
      company_name,
      nombre_tarifa AS llista_preu,
      DATE_TRUNC(data_alta, MONTH) AS activation_date,
      effective_dealer,
      affiliate_code,
      fecha_ultimo_cambio_tarifa,
      COUNT(*) AS EOPs
    FROM tarifa_llamada
    CROSS JOIN (
      SELECT DATE_SUB(dia, INTERVAL 1 DAY) AS fecha
      FROM (SELECT dia FROM UNNEST(GENERATE_DATE_ARRAY(DATE_SUB(DATE_TRUNC(CURRENT_DATE(), MONTH), INTERVAL 30 MONTH), CURRENT_DATE(), INTERVAL 1 MONTH)) AS dia)
    )
    WHERE data_alta <= fecha AND (data_baixa IS NULL OR data_baixa > fecha)
    GROUP BY ALL
),

eops_proporcional AS (
    SELECT
      CAST(CONCAT(SUBSTRING(yearmonth,1,4),'-',SUBSTRING(yearmonth,5,2), '-01') AS date) AS mes,
      activation_date,
      effective_dealer,
      affiliate_code,
      llista_preu AS tarifa,
      company_name AS tenant_code,
      fecha_ultimo_cambio_tarifa,
      EOPs,
      SUM(EOPs) OVER(PARTITION BY yearmonth, llista_preu, company_name) AS total_eops_tarifa
    FROM base_calculo
),

eops_cec_table AS (
    SELECT
      EXTRACT(YEAR FROM eop.mes) AS year,
      eop.activation_date,
      eop.effective_dealer,
      eop.affiliate_code,
      CASE
        WHEN STARTS_WITH(eop.affiliate_code, 'batman-30') THEN 'BATMAN 30€'
        WHEN STARTS_WITH(eop.affiliate_code, 'batman-50') THEN 'BATMAN 50€'
        WHEN STARTS_WITH(eop.affiliate_code, 'batman-40') THEN 'BATMAN 40€'
        WHEN STARTS_WITH(eop.affiliate_code, 'masmovil-bat-60') THEN 'BATMAN 60€'
        WHEN STARTS_WITH(eop.affiliate_code, 'energygo-bat-60') THEN 'BATMAN 60€'
        WHEN STARTS_WITH(eop.affiliate_code, 'descuento-captacion-30') THEN 'DTO CAPTA 30€'
        WHEN STARTS_WITH(eop.affiliate_code, 'descuento-captacion-50') THEN 'DTO CAPTA 50€'
        WHEN STARTS_WITH(eop.affiliate_code, 'energygo-bienvenida-30') THEN 'DTO CAPTA 30€'
        WHEN STARTS_WITH(eop.affiliate_code, 'masmovil-bienvenida-30') THEN 'DTO CAPTA 30€'
        WHEN STARTS_WITH(eop.affiliate_code, 'requete') THEN 'DTO CAPTA 30€'
        WHEN STARTS_WITH(eop.affiliate_code, 'masmovil-vuelta-al-cole') THEN 'DTO CAPTA 50€'
        WHEN STARTS_WITH(eop.affiliate_code, 'energygo-vuelta-al-cole') THEN 'DTO CAPTA 50€'
        ELSE NULL
      END AS descuento_capta,
      EXTRACT(MONTH FROM eop.mes) AS month,
      eop.tarifa,
      eop.tenant_code,
      eop.fecha_ultimo_cambio_tarifa,
      'EOP' AS kpi_type,
      eop.EOPs AS eops_brutos,
      SAFE_DIVIDE(eop.EOPs, eop.total_eops_tarifa) * COALESCE(altas.altas_cec_total, 0) AS altas_cec,
      SAFE_DIVIDE(eop.EOPs, eop.total_eops_tarifa) * COALESCE(bajas.bajas_cec_total, 0) AS bajas_cec,
      eop.EOPs
        - (SAFE_DIVIDE(eop.EOPs, eop.total_eops_tarifa) * COALESCE(altas.altas_cec_total, 0))
        + (SAFE_DIVIDE(eop.EOPs, eop.total_eops_tarifa) * COALESCE(bajas.bajas_cec_total, 0)) AS eops_cec
    FROM eops_proporcional eop
    LEFT JOIN (
        SELECT tenant_code, tarifa, DATE(year,month,1) AS yearmonth, SUM(clientes_efecto_consumo) AS altas_cec_total
        FROM clientes_efecto_consumo_table WHERE kpi_type = 'Altas' GROUP BY 1,2,3
    ) altas ON eop.tarifa = altas.tarifa AND eop.tenant_code = altas.tenant_code AND eop.mes = altas.yearmonth
    LEFT JOIN (
        SELECT tenant_code, tarifa, DATE(year,month,1) AS yearmonth, SUM(clientes_efecto_consumo) AS bajas_cec_total
        FROM clientes_efecto_consumo_table WHERE kpi_type = 'Bajas' GROUP BY 1,2,3
    ) bajas ON eop.tarifa = bajas.tarifa AND eop.tenant_code = bajas.tenant_code AND eop.mes = bajas.yearmonth
)

SELECT
    * EXCEPT(eops_brutos, eops_cec),
    GREATEST(eops_brutos, 0) AS eops_brutos,
    GREATEST(eops_cec, 0) AS eops_cec
FROM eops_cec_table
)
