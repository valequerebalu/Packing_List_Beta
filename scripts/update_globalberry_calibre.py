"""
Script para agregar value_map de calibres al mapping_config de Global Berry en Supabase.

Uso:
    python scripts/update_globalberry_calibre.py

Este script:
1. Conecta a Supabase
2. Obtiene el mapping_config actual de Global Berry
3. Identifica la columna de calibre
4. Agrega el transform value_map para convertir:
   - "12+" → "+12"
   - "14+" → "+14"  
   - "10+" → "+10"
5. Actualiza el registro en Supabase
"""

import json
import sys
from pathlib import Path

# Agregar el directorio raíz al path para importar db.database
sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import get_supabase


# Configuración de la transformación de calibres
CALIBRE_MAP = {
    "12+": "+12",
    "14+": "+14",
    "10+": "+10",
}

# Nombres posibles de la columna de calibre en el mapping_config
# (agregar más si es necesario)
POSSIBLE_CALIBRE_COLUMNS = [
    "CALIBRE",
    "calibre", 
    "Size",
    "size",
    "Tamaño",
    "tamaño",
    "CALIBRE_FRUTA",
    "calibre_fruta",
]


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
    
    # Buscar por nombres conocidos
    for col_name in POSSIBLE_CALIBRE_COLUMNS:
        if col_name in columns:
            return col_name
    
    # Buscar por contenido (si alguna columna tiene source = "packing" y field = "calibre")
    for col_name, col_config in columns.items():
        if isinstance(col_config, dict):
            if col_config.get("field") == "calibre":
                return col_name
            if col_config.get("source") == "packing" and "calibre" in col_config.get("field", "").lower():
                return col_name
    
    return None


def create_calibre_transform():
    """Crea la configuración del transform value_map para calibres."""
    return {
        "source": "transform",
        "transform": {
            "type": "value_map",
            "field": "calibre",
            "source": "packing",
            "map": CALIBRE_MAP,
            "default": None,
        },
    }


def update_template(template_id, mapping_config, calibre_column):
    """Actualiza el mapping_config del template en Supabase."""
    supabase = get_supabase()
    
    # Crear el nuevo mapping_config con el transform de calibres
    new_mapping_config = mapping_config.copy()
    new_mapping_config["columns"] = mapping_config["columns"].copy()
    new_mapping_config["columns"][calibre_column] = create_calibre_transform()
    
    # Actualizar en Supabase
    result = (
        supabase.table("templates")
        .update({"mapping_config": new_mapping_config})
        .eq("id", template_id)
        .execute()
    )
    
    return result.data


def main():
    """Función principal del script."""
    print("=== Actualización de calibres para Global Berry ===\n")
    
    # 1. Obtener template de Global Berry
    print("1. Buscando template de Global Berry...")
    template = get_template_globalberry()
    
    if not template:
        print("❌ No se encontró template activo para Global Berry")
        sys.exit(1)
    
    print(f"   Template encontrado: {template['name']} (ID: {template['id']})")
    
    # 2. Buscar columna de calibre
    print("\n2. Buscando columna de calibre en mapping_config...")
    calibre_column = find_calibre_column(template["mapping_config"])
    
    if not calibre_column:
        print("❌ No se encontró columna de calibre en el mapping_config")
        print("   Columnas disponibles:")
        for col in template["mapping_config"].get("columns", {}).keys():
            print(f"   - {col}")
        print("\n   Por favor, verifica el nombre de la columna en Supabase")
        sys.exit(1)
    
    print(f"   Columna encontrada: {calibre_column}")
    
    # 3. Mostrar configuración actual
    print("\n3. Configuración actual de la columna:")
    current_config = template["mapping_config"]["columns"][calibre_column]
    print(f"   {json.dumps(current_config, indent=2, ensure_ascii=False)}")
    
    # 4. Crear nueva configuración
    print("\n4. Nueva configuración que se aplicará:")
    new_config = create_calibre_transform()
    print(f"   {json.dumps(new_config, indent=2, ensure_ascii=False)}")
    
    # 5. Confirmar actualización
    print("\n5. ¿Desea aplicar esta actualización? (s/n): ", end="")
    respuesta = input().strip().lower()
    
    if respuesta != 's':
        print("   Actualización cancelada por el usuario")
        sys.exit(0)
    
    # 6. Actualizar template
    print("\n6. Actualizando template en Supabase...")
    try:
        update_template(template["id"], template["mapping_config"], calibre_column)
        print("   ✅ Template actualizado exitosamente")
    except Exception as e:
        print(f"   ❌ Error al actualizar template: {e}")
        sys.exit(1)
    
    # 7. Verificar actualización
    print("\n7. Verificando actualización...")
    updated_template = get_template_globalberry()
    
    if updated_template:
        updated_config = updated_template["mapping_config"]["columns"][calibre_column]
        print(f"   Configuración actualizada:")
        print(f"   {json.dumps(updated_config, indent=2, ensure_ascii=False)}")
        
        # Verificar que el map sea correcto
        if updated_config.get("transform", {}).get("map") == CALIBRE_MAP:
            print("\n✅ ¡Actualización completada exitosamente!")
            print("   Los calibres se mapearán correctamente:")
            print("   - '12+' → '+12'")
            print("   - '14+' → '+14'")
            print("   - '10+' → '+10'")
        else:
            print("\n⚠️  La actualización parece haber sido aplicada, pero el map no coincide")
    else:
        print("   ❌ No se pudo verificar la actualización")
    
    print("\n=== Proceso completado ===")


if __name__ == "__main__":
    main()
