# Configuración de Calibres para Global Berry

## Problema

Los calibres vienen en el archivo de Packing Beta como:
- `12+`
- `14+`
- `10+`

Pero la plantilla Excel de Global Berry espera:
- `+12`
- `+14`
- `+10`

## Solución

Configurar un `value_map` en el `mapping_config` de Global Berry en Supabase para mapear automáticamente los calibres.

## Scripts Disponibles

### 1. Verificar configuración actual
```bash
python scripts/verify_globalberry_config.py
```

Este script muestra:
- El template activo de Global Berry
- Las columnas del mapping_config
- La configuración de la columna de calibre
- Si el transform value_map está configurado correctamente

### 2. Actualizar configuración
```bash
python scripts/update_globalberry_calibre.py
```

Este script:
1. Conecta a Supabase
2. Obtiene el mapping_config actual de Global Berry
3. Identifica la columna de calibre
4. Agrega el transform value_map para convertir:
   - `"12+"` → `"+12"`
   - `"14+"` → `"+14"`  
   - `"10+"` → `"+10"`
5. Actualiza el registro en Supabase

### 3. Script SQL (alternativa)
```bash
# Ver el mapping_config actual
psql -f scripts/update_globalberry_calibre_v2.sql
```

## Transformación Configurada

El transform `value_map` se ve así en el mapping_config:

```json
{
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
}
```

## Cómo funciona

1. **Lectura**: El sistema lee el valor del campo `calibre` del archivo de Packing Beta
2. **Transformación**: El `value_map` busca el valor en el diccionario `map`
3. **Mapeo**: Si encuentra una coincidencia, retorna el valor mapeado
4. **Escritura**: El valor transformado se escribe en la plantilla Excel

## Notas Importantes

- El `value_map` es case-insensitive (no distingue mayúsculas/minúsculas)
- Si el valor no se encuentra en el `map`, se retorna el `default` (en este caso `null`)
- La configuración se almacena en la tabla `templates` de Supabase
- Los cambios se aplican inmediatamente a todas las generaciones futuras

## Troubleshooting

### Si no se encuentra la columna de calibre
1. Ejecutar `verify_globalberry_config.py` para ver las columnas disponibles
2. Verificar el nombre exacto de la columna en el mapping_config
3. Actualizar el script `update_globalberry_calibre.py` con el nombre correcto

### Si el transform no se aplica
1. Verificar que el `source` sea `"transform"`
2. Verificar que el `type` sea `"value_map"`
3. Verificar que el `field` sea `"calibre"`
4. Verificar que el `source` del transform sea `"packing"`

### Si los valores no se mapean correctamente
1. Verificar que los valores en el `map` coincidan exactamente con los del archivo de entrada
2. El `value_map` es case-insensitive, pero los valores deben coincidir en contenido
3. Verificar que no haya espacios extra en los valores
