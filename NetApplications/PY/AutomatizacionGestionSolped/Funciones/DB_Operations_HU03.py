# =========================================
# DB_Operations_HU03 - Operaciones Base de Datos HU03
# Autor: Senior Software Engineer
# Descripcion: Módulo para gestionar persistencia de validaciones
#              de SOLPEDs en SQL Server en lugar de archivos de texto/Excel
# Propiedad de Colsubsidio
# =========================================

from sqlalchemy import (
    text,
    Column,
    String,
    Integer,
    DateTime,
    Boolean,
    Numeric,
    ForeignKey,
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import datetime
import traceback
from EscribirLog import WriteLog
from Config.settings import RUTAS

Base = declarative_base()


def check_and_create_tables(engine, schema):
    """
    Verifica si las tablas de validación existen, si no las crea con el esquema completo.

    Args:
        engine: SQLAlchemy Engine de conexión a SQL Server
        schema: Nombre del schema (ej: "NotasCreditoYFacturacion")

    Returns:
        bool: True si operación exitosa, False si error
    """
    try:
        with engine.connect() as connection:
            # Crear tabla SOLPED_Validation si no existe
            create_solped_validation = f"""
            IF NOT EXISTS (
                SELECT TABLE_NAME 
                FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_SCHEMA = '{schema}' 
                AND TABLE_NAME = 'SOLPED_Validation'
            )
            BEGIN
                CREATE TABLE [{schema}].[SOLPED_Validation] (
                    [SOLPED_ID] NVARCHAR(50) NOT NULL PRIMARY KEY,
                    [Overall_Status] NVARCHAR(100) NOT NULL,
                    [Overall_Observations] NVARCHAR(MAX) NULL,
                    [Has_Attachments] BIT NULL,
                    [Attachment_Summary] NVARCHAR(MAX) NULL,
                    [Processing_Timestamp] DATETIME NOT NULL,
                    [Processing_User] NVARCHAR(100) NOT NULL,
                    [SAP_Error_Consult] NVARCHAR(MAX) NULL,
                    [CreatedAt] DATETIME DEFAULT GETDATE()
                )
            END
            """
            connection.execute(text(create_solped_validation))

            # Crear tabla SOLPED_Item_Validation si no existe
            create_solped_item_validation = f"""
            IF NOT EXISTS (
                SELECT TABLE_NAME 
                FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_SCHEMA = '{schema}' 
                AND TABLE_NAME = 'SOLPED_Item_Validation'
            )
            BEGIN
                CREATE TABLE [{schema}].[SOLPED_Item_Validation] (
                    [Item_Validation_ID] INT IDENTITY(1,1) PRIMARY KEY,
                    [SOLPED_ID] NVARCHAR(50) NOT NULL,
                    [Item_Number] NVARCHAR(50) NOT NULL,
                    [Item_Status] NVARCHAR(100) NOT NULL,
                    [Item_Observations] NVARCHAR(MAX) NULL,
                    [SAP_Item_Text] NVARCHAR(MAX) NULL,
                    
                    -- Campos extraídos de validación
                    [Extracted_Cantidad] NUMERIC(18,2) NULL,
                    [Extracted_Valor_Total] NUMERIC(18,2) NULL,
                    [Extracted_Valor_Unitario] NUMERIC(18,2) NULL,
                    [Extracted_Centro_Costo] NVARCHAR(100) NULL,
                    [Extracted_Cuenta_Contable] NVARCHAR(100) NULL,
                    
                    -- Resultados de validación
                    [Validation_Mandatory_Fields_Missing] NVARCHAR(MAX) NULL,
                    [Validation_Numeric_Errors] NVARCHAR(MAX) NULL,
                    [Validation_Format_Errors] NVARCHAR(MAX) NULL,
                    [Validation_Business_Rules] NVARCHAR(MAX) NULL,
                    
                    -- Timestamps y auditoría
                    [Processing_Timestamp] DATETIME NOT NULL,
                    [CreatedAt] DATETIME DEFAULT GETDATE(),
                    
                    -- Foreign Key
                    CONSTRAINT [FK_SOLPED_Item_Validation_SOLPED] 
                        FOREIGN KEY ([SOLPED_ID]) 
                        REFERENCES [{schema}].[SOLPED_Validation]([SOLPED_ID])
                        ON DELETE CASCADE
                )
            END
            """
            connection.execute(text(create_solped_item_validation))

            connection.commit()
            print("✅ Tablas de validación verificadas/creadas correctamente")
            return True

    except Exception as e:
        print(f"❌ Error al crear tablas: {e}")
        traceback.print_exc()
        return False


def insert_solped_validation(
    engine,
    schema,
    solped_id,
    overall_status,
    overall_observations,
    has_attachments=None,
    attachment_summary=None,
    processing_timestamp=None,
    processing_user="SYSTEM",
    sap_error_consult=None,
):
    """
    Inserta o actualiza el registro de validación principal de una SOLPED.
    Usa MERGE para manejar tanto inserciones como actualizaciones.

    Args:
        engine: SQLAlchemy Engine
        schema: Nombre del schema
        solped_id: ID de la SOLPED (ej: "1300139390")
        overall_status: Estado general (ej: "Con Adjuntos", "Sin Adjuntos", "Error", etc.)
        overall_observations: Observaciones generales
        has_attachments: Bool indicando si tiene adjuntos
        attachment_summary: Resumen de adjuntos
        processing_timestamp: Timestamp del procesamiento
        processing_user: Usuario que procesó (ej: "CGRPA009")
        sap_error_consult: Mensaje de error si lo hay

    Returns:
        bool: True si operación exitosa
    """
    try:
        if processing_timestamp is None:
            processing_timestamp = datetime.datetime.now()

        # Convertir tipos de datos correctamente
        has_attachments_bit = (
            1 if has_attachments else 0 if has_attachments is not None else None
        )

        merge_query = f"""
        MERGE INTO [{schema}].[SOLPED_Validation] AS target
        USING (
            SELECT 
                @solped_id AS SOLPED_ID,
                @overall_status AS Overall_Status,
                @overall_observations AS Overall_Observations,
                @has_attachments AS Has_Attachments,
                @attachment_summary AS Attachment_Summary,
                @processing_timestamp AS Processing_Timestamp,
                @processing_user AS Processing_User,
                @sap_error_consult AS SAP_Error_Consult
        ) AS source
        ON target.SOLPED_ID = source.SOLPED_ID
        WHEN MATCHED THEN
            UPDATE SET
                Overall_Status = source.Overall_Status,
                Overall_Observations = source.Overall_Observations,
                Has_Attachments = CASE WHEN source.Has_Attachments IS NOT NULL 
                                   THEN source.Has_Attachments 
                                   ELSE target.Has_Attachments END,
                Attachment_Summary = CASE WHEN source.Attachment_Summary IS NOT NULL
                                     THEN source.Attachment_Summary
                                     ELSE target.Attachment_Summary END,
                Processing_Timestamp = source.Processing_Timestamp,
                Processing_User = source.Processing_User,
                SAP_Error_Consult = CASE WHEN source.SAP_Error_Consult IS NOT NULL
                                   THEN source.SAP_Error_Consult
                                   ELSE target.SAP_Error_Consult END
        WHEN NOT MATCHED THEN
            INSERT (SOLPED_ID, Overall_Status, Overall_Observations, Has_Attachments, 
                    Attachment_Summary, Processing_Timestamp, Processing_User, SAP_Error_Consult)
            VALUES (source.SOLPED_ID, source.Overall_Status, source.Overall_Observations, 
                    source.Has_Attachments, source.Attachment_Summary, source.Processing_Timestamp, 
                    source.Processing_User, source.SAP_Error_Consult);
        """

        with engine.connect() as connection:
            connection.execute(
                text(merge_query),
                {
                    "solped_id": str(solped_id),
                    "overall_status": str(overall_status),
                    "overall_observations": (
                        str(overall_observations) if overall_observations else None
                    ),
                    "has_attachments": has_attachments_bit,
                    "attachment_summary": (
                        str(attachment_summary) if attachment_summary else None
                    ),
                    "processing_timestamp": processing_timestamp,
                    "processing_user": str(processing_user),
                    "sap_error_consult": (
                        str(sap_error_consult) if sap_error_consult else None
                    ),
                },
            )
            connection.commit()

        print(f"✅ SOLPED {solped_id} - Estado: {overall_status} (persistido en BD)")
        return True

    except Exception as e:
        print(f"❌ Error al insertar SOLPED_Validation para {solped_id}: {e}")
        traceback.print_exc()
        return False


def insert_solped_item_validation(
    engine,
    schema,
    solped_id,
    item_number,
    item_status,
    item_observations=None,
    sap_item_text=None,
    extracted_cantidad=None,
    extracted_valor_total=None,
    extracted_valor_unitario=None,
    extracted_centro_costo=None,
    extracted_cuenta_contable=None,
    validation_mandatory_fields_missing=None,
    validation_numeric_errors=None,
    validation_format_errors=None,
    validation_business_rules=None,
    processing_timestamp=None,
    processing_user="SYSTEM",
):
    """
    Inserta un registro de validación de item dentro de una SOLPED.

    Args:
        engine: SQLAlchemy Engine
        schema: Nombre del schema
        solped_id: ID de la SOLPED
        item_number: Número del item dentro de la SOLPED
        item_status: Estado del item (ej: "Validado", "Con Problemas", "Manual", etc.)
        item_observations: Observaciones del item
        sap_item_text: Texto del item extraído de SAP
        extracted_cantidad: Cantidad extraída
        extracted_valor_total: Valor total extraído
        extracted_valor_unitario: Valor unitario extraído
        extracted_centro_costo: Centro de costo extraído
        extracted_cuenta_contable: Cuenta contable extraída
        validation_mandatory_fields_missing: Campos obligatorios faltantes
        validation_numeric_errors: Errores numéricos encontrados
        validation_format_errors: Errores de formato encontrados
        validation_business_rules: Violaciones de reglas de negocio
        processing_timestamp: Timestamp del procesamiento
        processing_user: Usuario que procesó

    Returns:
        bool: True si operación exitosa
    """
    try:
        if processing_timestamp is None:
            processing_timestamp = datetime.datetime.now()

        # Convertir valores numéricos correctamente
        cantidad_numeric = None
        valor_total_numeric = None
        valor_unitario_numeric = None

        if extracted_cantidad is not None:
            try:
                cantidad_numeric = float(extracted_cantidad)
            except (ValueError, TypeError):
                cantidad_numeric = None

        if extracted_valor_total is not None:
            try:
                valor_total_numeric = float(extracted_valor_total)
            except (ValueError, TypeError):
                valor_total_numeric = None

        if extracted_valor_unitario is not None:
            try:
                valor_unitario_numeric = float(extracted_valor_unitario)
            except (ValueError, TypeError):
                valor_unitario_numeric = None

        insert_query = f"""
        INSERT INTO [{schema}].[SOLPED_Item_Validation] (
            SOLPED_ID,
            Item_Number,
            Item_Status,
            Item_Observations,
            SAP_Item_Text,
            Extracted_Cantidad,
            Extracted_Valor_Total,
            Extracted_Valor_Unitario,
            Extracted_Centro_Costo,
            Extracted_Cuenta_Contable,
            Validation_Mandatory_Fields_Missing,
            Validation_Numeric_Errors,
            Validation_Format_Errors,
            Validation_Business_Rules,
            Processing_Timestamp,
            CreatedAt
        )
        VALUES (
            @solped_id,
            @item_number,
            @item_status,
            @item_observations,
            @sap_item_text,
            @cantidad,
            @valor_total,
            @valor_unitario,
            @centro_costo,
            @cuenta_contable,
            @mandatory_fields,
            @numeric_errors,
            @format_errors,
            @business_rules,
            @processing_timestamp,
            GETDATE()
        )
        """

        with engine.connect() as connection:
            connection.execute(
                text(insert_query),
                {
                    "solped_id": str(solped_id),
                    "item_number": str(item_number),
                    "item_status": str(item_status),
                    "item_observations": (
                        str(item_observations) if item_observations else None
                    ),
                    "sap_item_text": str(sap_item_text) if sap_item_text else None,
                    "cantidad": cantidad_numeric,
                    "valor_total": valor_total_numeric,
                    "valor_unitario": valor_unitario_numeric,
                    "centro_costo": (
                        str(extracted_centro_costo) if extracted_centro_costo else None
                    ),
                    "cuenta_contable": (
                        str(extracted_cuenta_contable)
                        if extracted_cuenta_contable
                        else None
                    ),
                    "mandatory_fields": (
                        str(validation_mandatory_fields_missing)
                        if validation_mandatory_fields_missing
                        else None
                    ),
                    "numeric_errors": (
                        str(validation_numeric_errors)
                        if validation_numeric_errors
                        else None
                    ),
                    "format_errors": (
                        str(validation_format_errors)
                        if validation_format_errors
                        else None
                    ),
                    "business_rules": (
                        str(validation_business_rules)
                        if validation_business_rules
                        else None
                    ),
                    "processing_timestamp": processing_timestamp,
                },
            )
            connection.commit()

        print(f"   ✅ Item {item_number} de SOLPED {solped_id} registrado en BD")
        return True

    except Exception as e:
        print(
            f"❌ Error al insertar SOLPED_Item_Validation para {solped_id}/{item_number}: {e}"
        )
        traceback.print_exc()
        return False


def get_solped_validation_status(engine, schema, solped_id):
    """
    Obtiene el estado actual de una SOLPED desde la base de datos.

    Args:
        engine: SQLAlchemy Engine
        schema: Nombre del schema
        solped_id: ID de la SOLPED

    Returns:
        dict: Diccionario con los datos de la SOLPED, None si no existe
    """
    try:
        query = f"""
        SELECT 
            SOLPED_ID,
            Overall_Status,
            Overall_Observations,
            Has_Attachments,
            Attachment_Summary,
            Processing_Timestamp,
            Processing_User,
            SAP_Error_Consult,
            CreatedAt
        FROM [{schema}].[SOLPED_Validation]
        WHERE SOLPED_ID = @solped_id
        """

        with engine.connect() as connection:
            result = connection.execute(text(query), {"solped_id": str(solped_id)})
            row = result.fetchone()

            if row:
                return {
                    "solped_id": row[0],
                    "overall_status": row[1],
                    "overall_observations": row[2],
                    "has_attachments": row[3],
                    "attachment_summary": row[4],
                    "processing_timestamp": row[5],
                    "processing_user": row[6],
                    "sap_error_consult": row[7],
                    "created_at": row[8],
                }
            return None

    except Exception as e:
        print(f"❌ Error al consultar estado de SOLPED {solped_id}: {e}")
        return None


def get_solped_items_validation(engine, schema, solped_id):
    """
    Obtiene todos los items de validación de una SOLPED.

    Args:
        engine: SQLAlchemy Engine
        schema: Nombre del schema
        solped_id: ID de la SOLPED

    Returns:
        list: Lista de diccionarios con datos de items
    """
    try:
        query = f"""
        SELECT 
            Item_Validation_ID,
            Item_Number,
            Item_Status,
            Item_Observations,
            SAP_Item_Text,
            Extracted_Cantidad,
            Extracted_Valor_Total,
            Extracted_Centro_Costo,
            Processing_Timestamp,
            CreatedAt
        FROM [{schema}].[SOLPED_Item_Validation]
        WHERE SOLPED_ID = @solped_id
        ORDER BY Item_Number ASC
        """

        items = []
        with engine.connect() as connection:
            results = connection.execute(text(query), {"solped_id": str(solped_id)})

            for row in results:
                items.append(
                    {
                        "item_validation_id": row[0],
                        "item_number": row[1],
                        "item_status": row[2],
                        "item_observations": row[3],
                        "sap_item_text": row[4],
                        "extracted_cantidad": row[5],
                        "extracted_valor_total": row[6],
                        "extracted_centro_costo": row[7],
                        "processing_timestamp": row[8],
                        "created_at": row[9],
                    }
                )

        return items

    except Exception as e:
        print(f"❌ Error al consultar items de SOLPED {solped_id}: {e}")
        return []
