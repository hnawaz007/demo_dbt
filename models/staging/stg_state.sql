{{
  config(
    materialized='table'
  )
}}

SELECT *
FROM {{ source('src_sql', 'states') }}