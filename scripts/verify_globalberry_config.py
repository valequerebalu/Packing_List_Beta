"""
Script para verificar la configuración actual de Global Berry en Supabase.

Uso:
    python scripts/verify_globalberry_config.py

Este script muestra:
1. El template activo de Global Berry
2. Las columnas del mapping_config
3. La configuración de la columna de calibre
4. Si el transform value_map está configurado correctamente
"""

import json
import sys
from pathlib import Path

# Agregar el directorio raíz al path para importar db.database
sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import get_supabase


def get_template_globalberry():
    """Obtiene el template activo de Global Berry."""
    supabase = get_supabase()
    
    result = (
        supabase.table("templates")
        .select("id, name, mapping_config, customers(name), crops(name)")
        .eq("status", "active")
        .execute()
    )
    
    for row in result.data:
        customer_name = (row.get("customers") or {}).get("name", "")
        if customer_name == "Global Berry":
            mc = row["mapping_config"]
            if isinstance(mc, str):
                mc = json.loads(mc)
            row["mapping_config"] = mc
            return row
    
    return None


def find_calibre_column(mapping_config):
    """Busca la columna de calibre en el mapping_config."""
    columns = mapping_config.get("columns", {})
    
    # Nombres posibles de la columna de calibre
    possible_names = [
        "CALIBRE", "calibre", "Size", "size", "Tamaño", "tamaño",
        "CALIBRE_FRUTA", "calibre_fruta",
    ]
    
    # Buscar por nombres conocidos
    for col_name in possible_names:
        if col_name in columns:
            return col_name
    
    # Buscar por contenido
    for col_name, col_config in columns.items():
        if isinstance(col_config, dict):
            if col_config.get("field") == "calibre":
                return col_name
            if col_config.get("source") == "packing" and "calibre" in col_config.get("field", "").lower():
                return col_name
    
    return None


def main():
    """Función principal del script."""
    print("=== Verificación de configuración de Global Berry ===\n")
    
    # 1. Obtener template de Global Berry
    print("1. Buscando template de Global Berry...")
    template = get_template_globalberry()
    
    if not template:
        print("❌ No se encontró template activo para Global Berry")
        sys.exit(1)
    
    print(f"   Template encontrado: {template['name']} (ID: {template['id']})")
    print(f"   Cliente: {template['customers']['name']}")
    print(f"   Cultivo: {template['crops']['name']}")
    
    # 2. Mostrar todas las columnas del mapping_config
    print("\n2. Columnas del mapping_config:")
    columns = template["mapping_config"].get("columns", {})
    
    if not columns:
        print("   ⚠️  No hay columnas definidas en el mapping_config")
    else:
        for col_name, col_config in columns.items():
            print(f"\n   📋 {col_name}:")
            print(f"      {json.dumps(col_config, indent=6, ensure_ascii=False)}")
    
    # 3. Buscar columna de calibre
    print("\n3. Buscando columna de calibre...")
    calibre_column = find_calibre_column(template["mapping_config"])
    
    if not calibre_column:
        print("   ❌ No se encontró columna de calibre")
        print("   Esto significa que necesitas agregar la columna de calibre al mapping_config")
    else:
        print(f"   ✅ Columna de calibre encontrada: {calibre_column}")
        
        # 4. Verificar configuración de calibre
        print("\n4. Configuración de la columna de calibre:")
        calibre_config = template["mapping_config"]["columns"][calibre_column]
        print(f"   {json.dumps(calibre_config, indent=2, ensure_ascii=False)}")
        
        # 5. Verificar si tiene transform value_map
        print("\n5. Verificando transform value_map:")
        if calibre_config.get("source") == "transform":
            transform = calibre_config.get("transform", {})
            if transform.get("type") == "value_map":
                print("   ✅ Transform value_map encontrado")
                
                # Verificar el map
                calibre_map = transform.get("map", {})
                expected_map = {"12+": "+12", "14+": "+14", "10+": "+10"}
                
                if calibre_map == expected_map:
                    print("   ✅ Map de calibres configurado correctamente:")
                    print("      - '12+' → '+12'")
                    print("      - '14+' → '+14'")
                    print("      - '10+' → '+10'")
                else:
                    print("   ⚠️  Map de calibres no coincide con lo esperado:")
                    print(f"      Actual: {calibre_map}")
                    print(f"      Esperado: {expected_map}")
            else:
                print(f"   ❌ Transform type incorrecto: {transform.get('type')}")
                print(f"      Se esperaba: 'value_map'")
        else:
            print(f"   ❌ Source incorrecto: {calibre_config.get('source')}")
            print(f"      Se esperaba: 'transform'")
    
    # 6. Mostrar metadata del mapping_config
    print("\n6. Metadata del mapping_config:")
    metadata = template["mapping_config"].get("metadata", {})
    if metadata:
        print(f"   {json.dumps(metadata, indent=2, ensure_ascii=False)}")
    else:
        print("   No hay metadata definida")
    
    # 7. Mostrar otras configuraciones importantes
    print("\n7. Otras configuraciones:")
    print(f"   sheet_name: {template['mapping_config'].get('sheet_name')}")
    print(f"   header_row: {template['mapping_config'].get('header_row')}")
    print(f"   data_start_row: {template['mapping_config'].get('data_start_row')}")
    print(f"   input_header_row: {template['mapping_config'].get('input_header_row')}")
    
    print("\n=== Verificación completada ===")


if __name__ == "__main__":
    main()
