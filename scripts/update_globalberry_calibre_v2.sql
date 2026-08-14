-- Script para agregar value_map de calibres al mapping_config de Global Berry
-- Ejecutar en Supabase SQL Editor

-- PASO 1: Ver el mapping_config actual de Global Berry
-- Copiar el resultado y verificar la estructura de las columnas
SELECT 
    t.id as template_id,
    t.name as template_name,
    c.name as customer_name,
    cr.name as crop_name,
    t.mapping_config
FROM templates t
JOIN customers c ON t.customer_id = c.id
JOIN crops cr ON t.crop_id = cr.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';

-- PASO 2: Ver las columnas actuales del mapping_config
-- Esto muestra la estructura de 'columns' para entender qué nombre tiene la columna de calibre
SELECT 
    t.id,
    t.name,
    jsonb_object_keys(t.mapping_config->'columns') as column_name
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';

-- PASO 3: Agregar el transform value_map para calibres
-- IMPORTANTE: Ajustar 'CALIBRE' por el nombre real de la columna que aparece en el PASO 2
-- Si la columna se llama 'Size', 'Tamaño', etc., cambiarlo abajo

-- Opción A: Si la columna ya existe y tiene source = "packing", reemplazarla con el transform
UPDATE templates 
SET mapping_config = jsonb_set(
    mapping_config,
    '{columns,CALIBRE}',  -- CAMBIAR 'CALIBRE' por el nombre real de la columna
    '{
        "source": "transform",
        "transform": {
            "type": "value_map",
            "field": "calibre",
            "source": "packing",
            "map": {
                "12+": "+12",
                "14+": "+14",
                "10+": "+10"
            },
            "default": null
        }
    }'::jsonb
)
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active'
  AND templates.id = t.id;

-- Opción B: Si necesitas agregar una nueva columna (usar solo si la columna no existe)
-- Descomentar y ajustar el nombre de la columna de salida en la plantilla
/*
UPDATE templates 
SET mapping_config = jsonb_set(
    mapping_config,
    '{columns,TAMANIO}',  -- Nombre de la columna en la plantilla Excel
    '{
        "source": "transform",
        "transform": {
            "type": "value_map",
            "field": "calibre",
            "source": "packing",
            "map": {
                "12+": "+12",
                "14+": "+14",
                "10+": "+10"
            },
            "default": null
        }
    }'::jsonb
)
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active'
  AND templates.id = t.id;
*/

-- PASO 4: Verificar el cambio
SELECT 
    t.id,
    t.name,
    t.mapping_config->'columns'->'CALIBRE' as calibre_config  -- CAMBIAR 'CALIBRE' por el nombre real
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';

-- PASO 5: Verificar que el JSON es válido
SELECT 
    t.id,
    t.name,
    jsonb_typeof(t.mapping_config) as config_type,
    jsonb_object_keys(t.mapping_config->'columns') as column_name
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';
