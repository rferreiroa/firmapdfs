# Información sobre Plantillas Word

Este directorio contiene plantillas de ejemplo y documentación sobre cómo crear plantillas compatibles con FirmaPDFs.

## Estructura de una Plantilla Compatible

Para que FirmaPDFs pueda procesar correctamente tus documentos Word, deben seguir estas pautas:

### 1. Checkboxes (Casillas de verificación)

Hay dos formas de crear checkboxes en Word que FirmaPDFs puede detectar:

#### Opción A: Content Controls (Recomendado)

1. Ir a **Desarrollador** > **Controles** > **Casilla de verificación**
2. Agregar una etiqueta (Tag) al control:
   - Seleccionar el control
   - Ir a **Propiedades**
   - En **Tag** poner: `chkAprobado`, `chkRechazado`, etc.

#### Opción B: Campos de Formulario Legacy

1. Ir a **Desarrollador** > **Controles** > **Herramientas heredadas** > **Casilla de verificación**
2. Configurar el nombre del marcador

### 2. Campos de Fecha

Para campos de fecha que se rellenan automáticamente:

1. Insertar un **Control de contenido de texto**
2. En **Propiedades** > **Tag** poner: `Fecha` o `FechaDocumento`
3. O usar un campo de fecha de Content Control

### 3. Otros Campos

Para campos de texto como "Responsable", "Departamento", etc.:

1. Insertar **Control de contenido de texto**
2. En **Propiedades** > **Tag** poner el nombre del campo

## Nombres de Controles Reconocidos

FirmaPDFs busca estos nombres de controles (configurables en `rules.yaml`):

### Checkboxes
| Nombre Lógico | Tags Buscados |
|---------------|---------------|
| aprobado | chkAprobado, CheckBoxAprobado, Aprobado |
| rechazado | chkRechazado, CheckBoxRechazado, Rechazado |
| conforme | chkConforme, Conforme |
| no_conforme | chkNoConforme, No conforme |
| urgente | chkUrgente, Urgente |

### Campos
| Nombre Lógico | Tags Buscados |
|---------------|---------------|
| fecha | Fecha, DATE, FechaDocumento |
| responsable | Responsable, NombreResponsable |
| departamento | Departamento, Dept, Area |

## Ejemplo de Documento

Un documento típico tendría esta estructura:

```
┌─────────────────────────────────────────────────┐
│                 TÍTULO DEL DOCUMENTO             │
├─────────────────────────────────────────────────┤
│ Fecha: [Campo fecha]                            │
│ Referencia: [Campo referencia]                  │
├─────────────────────────────────────────────────┤
│                                                 │
│ [Contenido del documento...]                    │
│                                                 │
├─────────────────────────────────────────────────┤
│ RESOLUCIÓN:                                     │
│                                                 │
│ ☐ Aprobado    ☐ Rechazado    ☐ Pendiente       │
│                                                 │
│ ☐ Conforme    ☐ No Conforme                    │
│                                                 │
├─────────────────────────────────────────────────┤
│ Responsable: [Campo responsable]                │
│ Departamento: [Campo departamento]              │
│                                                 │
│ Firma: _____________________                    │
└─────────────────────────────────────────────────┘
```

## Crear Plantilla desde Cero

1. Abrir Word
2. Ir a **Archivo** > **Opciones** > **Personalizar cinta**
3. Activar la pestaña **Desarrollador**
4. Usar controles de contenido de la pestaña Desarrollador
5. Guardar como `.docx`

## Notas Importantes

- Los nombres de tags son **case-insensitive** (no distinguen mayúsculas)
- Si el nombre exacto no coincide, FirmaPDFs buscará coincidencias parciales
- Puedes añadir nuevos mappings en `config/rules.yaml`
- Los archivos temporales de Word (`~$*.docx`) son ignorados automáticamente
