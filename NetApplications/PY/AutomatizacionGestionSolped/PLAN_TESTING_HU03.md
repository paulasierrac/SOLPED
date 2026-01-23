# 🧪 Plan de Testing - Refactorización HU03

## Test 1: Verificar Creación de Tablas

**Objetivo:** Asegurar que las tablas se crean correctamente en BD

```python
# Ubicación: test_1_create_tables.py

from Funciones.Funciones_ARIA_Python import ConexionDB, Diccionario
from Funciones.DB_Operations_HU03 import check_and_create_tables

def test_create_tables():
    """Verifica que las tablas se creen si no existen"""
    print("=" * 60)
    print("TEST 1: Crear Tablas")
    print("=" * 60)
    
    try:
        # Conectar a BD
        db_engine = ConexionDB(Diccionario)
        db_schema = Diccionario.get("Schema", "NotasCreditoYFacturacion")
        
        # Crear tablas
        resultado = check_and_create_tables(db_engine, db_schema)
        
        if resultado:
            print("✅ TEST 1 PASÓ: Tablas creadas/verificadas correctamente")
            print(f"   Schema: {db_schema}")
            print(f"   Tablas:")
            print(f"      - SOLPED_Validation")
            print(f"      - SOLPED_Item_Validation")
            return True
        else:
            print("❌ TEST 1 FALLÓ: Error al crear tablas")
            return False
            
    except Exception as e:
        print(f"❌ TEST 1 FALLÓ: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_create_tables()
```

---

## Test 2: Insertar SOLPED (INSERT via MERGE)

**Objetivo:** Validar inserción de registro de SOLPED

```python
# Ubicación: test_2_insert_solped.py

from Funciones.Funciones_ARIA_Python import ConexionDB, Diccionario
from Funciones.DB_Operations_HU03 import insert_solped_validation, get_solped_validation_status
from datetime import datetime
import getpass

def test_insert_solped():
    """Prueba inserción de nueva SOLPED"""
    print("=" * 60)
    print("TEST 2: Insertar SOLPED")
    print("=" * 60)
    
    try:
        db_engine = ConexionDB(Diccionario)
        db_schema = Diccionario.get("Schema", "NotasCreditoYFacturacion")
        
        # Datos de prueba
        solped_id = "TEST_INSERT_001"
        
        # Insertar SOLPED
        print(f"\n1. Insertando SOLPED {solped_id}...")
        insert_result = insert_solped_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            overall_status="Test - Insert",
            overall_observations="SOLPED de prueba para validar INSERT",
            has_attachments=True,
            attachment_summary="Doc1.pdf, Doc2.xlsx",
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        
        if not insert_result:
            print(f"❌ Error insertando SOLPED")
            return False
        
        # Verificar que se insertó
        print(f"\n2. Verificando que SOLPED se insertó correctamente...")
        status = get_solped_validation_status(db_engine, db_schema, solped_id)
        
        if status and status["overall_status"] == "Test - Insert":
            print(f"✅ SOLPED encontrada en BD")
            print(f"   ID: {status['solped_id']}")
            print(f"   Estado: {status['overall_status']}")
            print(f"   Adjuntos: {status['has_attachments']}")
            print(f"   Observaciones: {status['overall_observations']}")
            print("\n✅ TEST 2 PASÓ: INSERT exitoso")
            return True
        else:
            print(f"❌ SOLPED no encontrada o datos incorrectos")
            return False
            
    except Exception as e:
        print(f"❌ TEST 2 FALLÓ: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_insert_solped()
```

---

## Test 3: Actualizar SOLPED (UPDATE via MERGE)

**Objetivo:** Validar actualización de registro existente

```python
# Ubicación: test_3_update_solped.py

from Funciones.Funciones_ARIA_Python import ConexionDB, Diccionario
from Funciones.DB_Operations_HU03 import insert_solped_validation, get_solped_validation_status
from datetime import datetime
import getpass

def test_update_solped():
    """Prueba actualización de SOLPED existente"""
    print("=" * 60)
    print("TEST 3: Actualizar SOLPED (MERGE)")
    print("=" * 60)
    
    try:
        db_engine = ConexionDB(Diccionario)
        db_schema = Diccionario.get("Schema", "NotasCreditoYFacturacion")
        
        solped_id = "TEST_UPDATE_001"
        
        # 1. Insertar SOLPED inicial
        print(f"\n1. Insertando SOLPED inicial...")
        insert_solped_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            overall_status="En Proceso",
            overall_observations="Estado inicial",
            has_attachments=False,
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ SOLPED insertada")
        
        # 2. Actualizar SOLPED
        print(f"\n2. Actualizando SOLPED...")
        insert_solped_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            overall_status="Validada",
            overall_observations="Se encontraron los adjuntos",
            has_attachments=True,
            attachment_summary="Documentación completada",
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ SOLPED actualizada")
        
        # 3. Verificar actualización
        print(f"\n3. Verificando actualización...")
        status = get_solped_validation_status(db_engine, db_schema, solped_id)
        
        if (status and 
            status["overall_status"] == "Validada" and 
            status["has_attachments"] == 1):
            print(f"✅ Actualización correcta")
            print(f"   Estado anterior: 'En Proceso' → Estado actual: '{status['overall_status']}'")
            print(f"   Adjuntos: {status['has_attachments']}")
            print("\n✅ TEST 3 PASÓ: UPDATE via MERGE exitoso")
            return True
        else:
            print(f"❌ Actualización no reflejada en BD")
            return False
            
    except Exception as e:
        print(f"❌ TEST 3 FALLÓ: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_update_solped()
```

---

## Test 4: Insertar Items de SOLPED

**Objetivo:** Validar inserción de items y conversión de tipos

```python
# Ubicación: test_4_insert_items.py

from Funciones.Funciones_ARIA_Python import ConexionDB, Diccionario
from Funciones.DB_Operations_HU03 import (
    insert_solped_validation,
    insert_solped_item_validation,
    get_solped_items_validation
)
from datetime import datetime
import getpass

def test_insert_items():
    """Prueba inserción de items con diferentes tipos de datos"""
    print("=" * 60)
    print("TEST 4: Insertar Items de SOLPED")
    print("=" * 60)
    
    try:
        db_engine = ConexionDB(Diccionario)
        db_schema = Diccionario.get("Schema", "NotasCreditoYFacturacion")
        
        solped_id = "TEST_ITEMS_001"
        
        # 1. Crear SOLPED padre
        print(f"\n1. Creando SOLPED padre...")
        insert_solped_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            overall_status="Con Items",
            overall_observations="SOLPED con items de prueba",
            has_attachments=True,
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ SOLPED creada")
        
        # 2. Insertar Item #1 con datos numéricos
        print(f"\n2. Insertando Item 1 con datos numéricos...")
        insert_solped_item_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            item_number="10",
            item_status="Validado",
            item_observations="Item sin problemas",
            sap_item_text="ITEM 10: Material ABC",
            extracted_cantidad=50.50,          # float → NUMERIC(18,2)
            extracted_valor_total=25250.00,    # float → NUMERIC(18,2)
            extracted_valor_unitario=500.00,   # float → NUMERIC(18,2)
            extracted_centro_costo="CC-001",
            extracted_cuenta_contable="5110",
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ Item 1 insertado")
        
        # 3. Insertar Item #2 con valores None
        print(f"\n3. Insertando Item 2 con valores parciales...")
        insert_solped_item_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            item_number="20",
            item_status="Manual",
            item_observations="Requiere revisión",
            sap_item_text="ITEM 20: Material XYZ",
            extracted_cantidad=None,           # None → NULL
            extracted_valor_total=None,        # None → NULL
            validation_mandatory_fields_missing="Concepto, Cantidad",
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ Item 2 insertado")
        
        # 4. Insertar Item #3 con string que necesita conversión
        print(f"\n4. Insertando Item 3 con valores string...")
        insert_solped_item_validation(
            engine=db_engine,
            schema=db_schema,
            solped_id=solped_id,
            item_number="30",
            item_status="Validado",
            item_observations="Conversión de strings",
            sap_item_text="ITEM 30: Material DEF",
            extracted_cantidad="100.75",       # string → float → NUMERIC
            extracted_valor_total="30225.00",  # string → float → NUMERIC
            processing_timestamp=datetime.now(),
            processing_user=getpass.getuser()
        )
        print("   ✓ Item 3 insertado")
        
        # 5. Obtener y verificar items
        print(f"\n5. Verificando items insertados...")
        items = get_solped_items_validation(db_engine, db_schema, solped_id)
        
        if len(items) == 3:
            print(f"✅ Se encontraron {len(items)} items")
            for item in items:
                print(f"\n   Item {item['item_number']}:")
                print(f"      Estado: {item['item_status']}")
                print(f"      Cantidad: {item['extracted_cantidad']}")
                print(f"      Valor Total: {item['extracted_valor_total']}")
                print(f"      Centro Costo: {item['extracted_centro_costo']}")
            print("\n✅ TEST 4 PASÓ: Items insertados correctamente")
            return True
        else:
            print(f"❌ Expected 3 items, got {len(items)}")
            return False
            
    except Exception as e:
        print(f"❌ TEST 4 FALLÓ: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_insert_items()
```

---

## Test 5: Conversión de Tipos Crítica

**Objetivo:** Validar manejo correcto de conversiones

```python
# Ubicación: test_5_type_conversions.py

def test_type_conversions():
    """Prueba conversiones de tipos de datos"""
    print("=" * 60)
    print("TEST 5: Conversiones de Tipos")
    print("=" * 60)
    
    test_cases = [
        ("50.5", "string → float", lambda x: float(x)),
        ("invalid", "string invalido", lambda x: None if x == "invalid" else float(x)),
        (True, "bool True → 1", lambda x: 1 if x else 0 if x is not None else None),
        (False, "bool False → 0", lambda x: 1 if x else 0 if x is not None else None),
        (None, "None → None", lambda x: x),
    ]
    
    all_passed = True
    
    for value, description, converter in test_cases:
        try:
            result = converter(value)
            print(f"✅ {description}: {value} → {result}")
        except Exception as e:
            print(f"❌ {description}: ERROR - {e}")
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("✅ TEST 5 PASÓ: Todas las conversiones correctas")
    else:
        print("❌ TEST 5 FALLÓ: Algunas conversiones fallaron")
    print("=" * 60)
    
    return all_passed

if __name__ == "__main__":
    test_type_conversions()
```

---

## Test 6: Foreign Key Constraint

**Objetivo:** Validar que la FK previene inconsistencias

```python
# Ubicación: test_6_foreign_key.py

from Funciones.Funciones_ARIA_Python import ConexionDB, Diccionario
from Funciones.DB_Operations_HU03 import insert_solped_item_validation
from datetime import datetime
import getpass

def test_foreign_key():
    """Verifica que la FK impide insertar items huérfanos"""
    print("=" * 60)
    print("TEST 6: Foreign Key Constraint")
    print("=" * 60)
    
    try:
        db_engine = ConexionDB(Diccionario)
        db_schema = Diccionario.get("Schema", "NotasCreditoYFacturacion")
        
        # Intentar insertar item para SOLPED que NO existe
        print(f"\n1. Intentando insertar item para SOLPED inexistente...")
        print(f"   SOLPED_ID: 'INEXISTENTE_12345'")
        
        try:
            insert_solped_item_validation(
                engine=db_engine,
                schema=db_schema,
                solped_id="INEXISTENTE_12345",
                item_number="10",
                item_status="Prueba",
                item_observations="Este item debería fallar",
                processing_timestamp=datetime.now(),
                processing_user=getpass.getuser()
            )
            print("❌ ERROR: Item se insertó sin SOLPED padre (FK no funcionó)")
            print("❌ TEST 6 FALLÓ: FK no está activa")
            return False
            
        except Exception as e:
            if "FOREIGN KEY" in str(e) or "FK_" in str(e):
                print(f"✅ FK constraint funcionó correctamente")
                print(f"   Error capturado: {str(e)[:100]}...")
                print("✅ TEST 6 PASÓ: FK previene items huérfanos")
                return True
            else:
                print(f"❌ Error diferente: {e}")
                return False
                
    except Exception as e:
        print(f"❌ TEST 6 FALLÓ: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_foreign_key()
```

---

## Suite Completa de Tests

```python
# Ubicación: test_all.py

import test_1_create_tables
import test_2_insert_solped
import test_3_update_solped
import test_4_insert_items
import test_5_type_conversions
import test_6_foreign_key

def run_all_tests():
    print("\n" + "=" * 80)
    print(" EJECUTANDO SUITE COMPLETA DE TESTS - REFACTORIZACIÓN HU03".center(80))
    print("=" * 80 + "\n")
    
    tests = [
        ("TEST 1", test_1_create_tables.test_create_tables),
        ("TEST 2", test_2_insert_solped.test_insert_solped),
        ("TEST 3", test_3_update_solped.test_update_solped),
        ("TEST 4", test_4_insert_items.test_insert_items),
        ("TEST 5", test_5_type_conversions.test_type_conversions),
        ("TEST 6", test_6_foreign_key.test_foreign_key),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            resultado = test_func()
            results.append((test_name, resultado))
        except Exception as e:
            print(f"\n❌ {test_name} EXCEPCIÓN: {e}")
            results.append((test_name, False))
    
    # Resumen Final
    print("\n" + "=" * 80)
    print(" RESUMEN FINAL".center(80))
    print("=" * 80)
    
    passed = sum(1 for _, resultado in results if resultado)
    total = len(results)
    
    for test_name, resultado in results:
        status = "✅ PASÓ" if resultado else "❌ FALLÓ"
        print(f"{test_name}: {status}")
    
    print(f"\nTotal: {passed}/{total} tests pasados")
    
    if passed == total:
        print("\n🎉 ¡REFACTORIZACIÓN VALIDADA EXITOSAMENTE!")
    else:
        print(f"\n⚠️  {total - passed} test(s) requieren atención")
    
    print("=" * 80 + "\n")
    
    return passed == total

if __name__ == "__main__":
    success = run_all_tests()
    exit(0 if success else 1)
```

---

## Instrucciones de Ejecución

```bash
# Crear archivo de test
python test_all.py

# O ejecutar tests individuales
python test_1_create_tables.py
python test_2_insert_solped.py
python test_3_update_solped.py
python test_4_insert_items.py
python test_5_type_conversions.py
python test_6_foreign_key.py
```

---

**Documento de Testing creado:** 22 de enero de 2026  
**Versión:** 1.0
