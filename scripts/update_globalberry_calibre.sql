-- Script para agregar value_map de calibres al mapping_config de Global Berry
-- Ejecutar en Supabase SQL Editor

-- Primero, ver el mapping_config actual de Global Berry
SELECT 
    t.id,
    t.name,
    t.mapping_config
FROM templates t
JOIN customers c ON t.customer_id = c.id
JOIN crops cr ON t.crop_id = cr.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';

-- Actualizar el mapping_config para agregar el transform de calibres
-- Este script asume que la columna de calibre en la plantilla se llama "CALIBRE" o similar
-- y que el campo de origen en el packing se llama "calibre"

-- Opción 1: Si ya existe una columna de calibre en mapping_config, agregarle el transform
-- Opción 2: Si no existe, agregar una nueva columna con el transform

-- Script para actualizar (ajustar el nombre de la columna según el mapping_config actual):
UPDATE templates 
SET mapping_config = jsonb_set(
    mapping_config,
    '{columns,CALIBRE}',  -- Ajustar el nombre de la columna según el mapping_config actual
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

-- Verificar el cambio
SELECT 
    t.id,
    t.name,
    t.mapping_config->'columns'->'CALIBRE' as calibre_config
FROM templates t
JOIN customers c ON t.customer_id = c.id
WHERE c.name = 'Global Berry'
  AND t.status = 'active';
