with source as (

    select * from {{ source('src_postgres', 'src_dimproductcategory') }}
),
renamed as (

    select 
		productcategorykey, 
		frenchproductcategoryname, 
		englishproductcategoryname, 
		spanishproductcategoryname, 
		productcategoryalternatekey
    from source
)

select * from renamed
