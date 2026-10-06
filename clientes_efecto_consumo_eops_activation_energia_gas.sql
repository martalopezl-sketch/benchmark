CREATE or REPLACE table `mm-new-business-reporting.ENERGY.clientes_efecto_consumo_eops_activation_energia_gas` AS (

WITH clientes_efecto_consumo_table as (
    SELECT * FROM `mm-new-business-reporting.ENERGY.clientes_efecto_consumo_altas_bajas_energia_gas`
),

-- Código de tarifa -> nombre (el CEC y el Data Studio trabajan con el nombre)
pricelist_names AS (
    SELECT code, MIN(name) AS name
    FROM `mm-datamart-kd.ENERGY.lucera_product_pricelist`
    GROUP BY code
),

-- Cada fila de sales_energy_mo es un tramo de tarifa de un contrato (sin solapes).
-- Solo gas (G_SALE): se excluyen luz (E_*) y servicios (E_SVA / G_SVA).
-- Las filas change_type = 'change' son los cambios de tarifa del contrato.
tarifa_llamada AS (
    SELECT DISTINCT
        s.signup_id,
        s.terminal_contract_id AS polissa_id,
        s.energy_effective_dealer AS effective_dealer,
        DATE(s.registered_activation_date) AS data_alta,
        DATE(s.registered_deactivation_date) AS data_baixa,
        IF(s.change_type = 'change', DATE(s.registered_activation_date), NULL) AS fecha_ultimo_cambio_tarifa,
        CASE WHEN COALESCE(p.name, s.energy_pricelist_code) IN ('Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)', 'Fijo 24H Básica Retención Natural (Abril 2023 - 0,1288)')
             THEN 'Fijo 24H Básica Retención Natural (Marzo 2023 - 0,1288)'
             ELSE COALESCE(p.name, s.energy_pricelist_code) END AS nombre_tarifa,
        s.energy_brand AS company_name,
        s.access_tariff_name AS tarifa_acceso
    FROM `mm-datamart-kd.ENERGY.sales_energy_mo` s
    LEFT JOIN pricelist_names p ON p.code = s.energy_pricelist_code
    WHERE s.energy_pricelist_code LIKE 'G_SALE%'
      AND s.registered_activation_date IS NOT NULL
      AND DATE(s.registered_activation_date) <= DATE_SUB(CURRENT_DATE(), INTERVAL 1 DAY)
),

base AS (
    SELECT
        FORMAT_DATE('%Y%m',fecha) AS yearmonth,
        'Gas' AS producto,
        company_name,
        nombre_tarifa as llista_preu,
        DATE_TRUNC(data_alta, MONTH) as activation_date,
        effective_dealer,
        tarifa_acceso,
        fecha_ultimo_cambio_tarifa
    FROM tarifa_llamada
    CROSS JOIN (
        SELECT DATE_SUB(dia, INTERVAL 1 day) AS fecha
        FROM (SELECT dia FROM UNNEST(GENERATE_DATE_ARRAY(DATE_SUB(DATE_TRUNC(CURRENT_DATE(),month), INTERVAL 30 MONTH), CURRENT_DATE(),INTERVAL 1 MONTH)) AS dia)
    )
    WHERE data_alta <= fecha
      AND (data_baixa IS NULL OR data_baixa > fecha)
),

bases AS (
    SELECT
        yearmonth,
        company_name,
        llista_preu,
        producto,
        activation_date,
        effective_dealer,
        tarifa_acceso,
        fecha_ultimo_cambio_tarifa,
        COUNTIF(company_name = 'lucera') AS non_mass_base_luc,
        COUNTIF(company_name IN ('energygo', 'yoigo')) AS non_mass_base_ego,
        COUNTIF(company_name = 'masmovil') AS non_mass_base_mme,
        COUNTIF(company_name = 'euskaltel') AS non_mass_base_eusk,
        COUNTIF(company_name = 'r') AS non_mass_base_r,
        COUNTIF(company_name = 'pepe') AS non_mass_base_pepe,
        COUNTIF(company_name = 'orange') AS non_mass_base_orange,
        COUNTIF(company_name = 'jazztel') AS non_mass_base_jazztel
    FROM base
    GROUP BY ALL
),

eops_mes as (
    SELECT
        mes,
        activation_date,
        effective_dealer,
        tarifa_acceso,
        llista_preu as current_lucera_tariff_name,
        company_name as tenant_code,
        producto,
        fecha_ultimo_cambio_tarifa,
        sum(total_eops) as EOPs
    FROM (
        SELECT
            cast(concat(substring(yearmonth,1,4),'-',substring(yearmonth,5,2), '-01') as date) as mes,
            company_name,
            llista_preu,
            producto,
            effective_dealer,
            tarifa_acceso,
            fecha_ultimo_cambio_tarifa,
            (non_mass_base_luc + non_mass_base_ego + non_mass_base_mme + non_mass_base_pepe + non_mass_base_eusk + non_mass_base_r + non_mass_base_orange + non_mass_base_jazztel) as total_eops,
            'Energía colectiva' AS fuente,
            activation_date
        FROM bases

        UNION ALL

        -- Foto del mes en curso: contratos de gas activos hoy
        SELECT
            cast(DATE_TRUNC(CURRENT_DATE(), MONTH) as date) AS mes,
            t.company_name,
            t.nombre_tarifa as llista_preu,
            'Gas' AS producto,
            t.effective_dealer,
            t.tarifa_acceso,
            NULL AS fecha_ultimo_cambio_tarifa,  -- No disponible en esta rama del UNION
            COUNTIF(t.data_baixa IS NULL) as total_eops,
            'Energía colectiva' AS fuente,
            t.data_alta as activation_date
        FROM tarifa_llamada t
        GROUP BY ALL
    ) result
    GROUP BY ALL
),

eops_cec_table as (
    SELECT
        EXTRACT(YEAR FROM eop.mes) as year,
        eop.activation_date,
        eop.effective_dealer,
        eop.tarifa_acceso,
        EXTRACT(MONTH FROM eop.mes) as month,
        eop.current_lucera_tariff_name as tarifa,
        eop.tenant_code,
        eop.fecha_ultimo_cambio_tarifa,
        'EOP' as kpi_type,
        CASE
            WHEN DATE_TRUNC(eop.mes, MONTH) = DATE_TRUNC(CURRENT_DATE(), MONTH)
            THEN eop.eops + IFNULL(altas_ayer.contracts, 0) - IFNULL(bajas_ayer.contracts, 0)
            ELSE eop.eops
        END as eops_brutos,
        SAFE_DIVIDE(eop.eops, SUM(eop.eops) OVER(PARTITION BY eop.mes, eop.current_lucera_tariff_name, eop.tenant_code, eop.activation_date)) as peso_acceso,
        altas.contracts as total_contracts_activated,
        bajas.contracts as total_contracts_churned,
        altas.clientes_efecto_consumo as total_altas_cec,
        bajas.clientes_efecto_consumo as total_bajas_cec
    FROM (
        SELECT mes, activation_date, effective_dealer, tarifa_acceso, current_lucera_tariff_name, tenant_code, fecha_ultimo_cambio_tarifa, EOPs as eops
        FROM eops_mes
    ) eop
    LEFT JOIN (SELECT tenant_code, tarifa, date(year,month,1) as yearmonth, sum(clientes_efecto_consumo) as clientes_efecto_consumo, sum(contracts) as contracts FROM clientes_efecto_consumo_table where kpi_type = 'Altas' group by 1,2,3) altas
        ON eop.current_lucera_tariff_name = altas.tarifa AND eop.tenant_code = altas.tenant_code AND eop.mes = altas.yearmonth
    LEFT JOIN (SELECT tenant_code, tarifa, date(year,month,1) as yearmonth, sum(contracts) as contracts FROM clientes_efecto_consumo_table where kpi_type = 'Altas' AND DATE(year, month, day) >= CURRENT_DATE() group by all) altas_ayer
        ON eop.current_lucera_tariff_name = altas_ayer.tarifa AND eop.tenant_code = altas_ayer.tenant_code AND eop.mes = altas_ayer.yearmonth
    LEFT JOIN (SELECT tenant_code, tarifa, date(year,month,1) as yearmonth, sum(clientes_efecto_consumo) as clientes_efecto_consumo, sum(contracts) as contracts FROM clientes_efecto_consumo_table where kpi_type = 'Bajas' group by 1,2,3) bajas
        ON eop.current_lucera_tariff_name = bajas.tarifa AND eop.tenant_code = bajas.tenant_code AND eop.mes = bajas.yearmonth
    LEFT JOIN (SELECT tenant_code, tarifa, date(year,month,1) as yearmonth, sum(contracts) as contracts FROM clientes_efecto_consumo_table where kpi_type = 'Bajas' AND DATE(year, month, day) >= CURRENT_DATE() group by all) bajas_ayer
        ON eop.current_lucera_tariff_name = bajas_ayer.tarifa AND eop.tenant_code = bajas_ayer.tenant_code AND eop.mes = bajas_ayer.yearmonth
)

SELECT
    year,
    activation_date,
    effective_dealer,
    tarifa_acceso,
    month,
    tarifa,
    tenant_code,
    fecha_ultimo_cambio_tarifa,
    kpi_type,
    COALESCE(total_contracts_activated, 0) * peso_acceso as contracts_activated,
    COALESCE(total_contracts_churned, 0) * peso_acceso as contracts_churned,
    COALESCE(total_altas_cec, 0) * peso_acceso as altas_cec,
    COALESCE(total_bajas_cec, 0) * peso_acceso as bajas_cec,
    CASE WHEN eops_brutos < 0 THEN 0 ELSE eops_brutos END as eops_brutos,
    CASE
        WHEN (eops_brutos - (COALESCE(total_altas_cec, 0) * peso_acceso) + (COALESCE(total_bajas_cec, 0) * peso_acceso)) < 0 THEN 0
        ELSE (eops_brutos - (COALESCE(total_altas_cec, 0) * peso_acceso) + (COALESCE(total_bajas_cec, 0) * peso_acceso))
    END as eops_cec
FROM eops_cec_table
)
