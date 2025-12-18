# FirmaPDFs - Sistema de Automatización Documental

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Windows](https://img.shields.io/badge/platform-Windows%2011-lightgrey.svg)](https://www.microsoft.com/windows)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Sistema automatizado para procesar documentos Word, exportarlos a PDF y firmarlos digitalmente con certificado .pfx.

## 🎯 Características

- ✅ **Relleno automático de documentos Word** (checkboxes, fechas, campos)
- ✅ **Exportación a PDF** con máxima fidelidad
- ✅ **Firma digital** con certificado .pfx
- ✅ **Firma visible** (sello con nombre y fecha)
- ✅ **Reglas configurables** vía YAML
- ✅ **Detección flexible** de controles aunque varíe el texto
- ✅ **Report CSV** de resultados
- ✅ **Modo dry-run** para simulación
- ✅ **Idempotencia** (no procesa dos veces el mismo archivo)
- ✅ **Modo watch** para monitorizar carpeta
- ✅ **Almacenamiento seguro** de contraseñas

## 📋 Requisitos

- **Windows 11** (o Windows 10)
- **Python 3.8+** ([Descargar](https://www.python.org/downloads/))
- **Microsoft Word** instalado (para procesamiento COM)
- **Certificado digital** (.pfx o .p12)

## 🚀 Instalación Rápida

### 1. Clonar o descargar el proyecto

```batch
git clone https://github.com/example/firmapdfs.git
cd firmapdfs
```

### 2. Ejecutar instalador

```batch
install.bat
```

Esto creará un entorno virtual y instalará todas las dependencias.

### 3. Activar entorno virtual

```batch
venv\Scripts\activate.bat
```

### 4. Verificar instalación

```batch
firmapdfs --version
firmapdfs status
```

## 📁 Estructura del Proyecto

```
firmapdfs/
├── config/
│   ├── config.yaml          # Configuración principal
│   ├── rules.yaml            # Reglas de procesamiento
│   └── certificate.pfx       # Tu certificado (añadir manualmente)
├── samples/
│   ├── inbox/                # Documentos a procesar
│   └── templates/            # Plantillas de ejemplo
├── processed/
│   ├── docx/                 # Documentos Word procesados
│   ├── pdf/                  # PDFs exportados
│   ├── signed/               # PDFs firmados (salida final)
│   └── archive/              # Archivos procesados (opcional)
├── logs/
│   ├── firmapdfs_*.log       # Logs diarios
│   └── report.csv            # Reporte de procesamiento
├── src/firmapdfs/            # Código fuente
├── tests/                    # Tests
├── install.bat               # Instalador
├── run.bat                   # Ejecutar CLI
├── process.bat               # Procesar documentos
├── watch.bat                 # Modo vigilancia
└── README.md
```

## ⚙️ Configuración

### Configuración Principal (`config/config.yaml`)

```yaml
# Directorios
paths:
  inbox: "samples/inbox"
  output_signed: "processed/signed"

# Procesamiento Word
word:
  engine: "com"              # "com" (recomendado) o "docx"
  date_format: "%d/%m/%Y"

# Firma digital
signature:
  certificate_path: "config/certificate.pfx"
  password_storage: "credential_manager"  # Recomendado

  appearance:
    visible: true
    position: "bottom_right"
    text_template: |
      Firmado digitalmente por:
      {signer_name}
      Fecha: {date}
```

### Configurar Contraseña del Certificado

**Opción 1: Windows Credential Manager (Recomendado)**

```batch
firmapdfs set-password --method credential_manager
```

**Opción 2: Variable de entorno**

```batch
set FIRMAPDFS_CERT_PASSWORD=tu_contraseña
```

**Opción 3: En config.yaml (NO recomendado para producción)**

```yaml
signature:
  password_storage: "config"
  password: "tu_contraseña"  # ⚠️ Inseguro
```

### Reglas de Procesamiento (`config/rules.yaml`)

Define qué checkboxes marcar según el contenido del documento:

```yaml
rules:
  - name: "approval_documents"
    description: "Marcar aprobado si contiene texto de aprobación"
    conditions:
      - type: "contains_any"
        values:
          - "documento aprobado"
          - "se aprueba"
    actions:
      - type: "set_checkbox"
        checkbox: "aprobado"
        value: true
```

## 🖥️ Uso

### Procesar todos los documentos

```batch
firmapdfs process
```

### Procesar un archivo específico

```batch
firmapdfs process -f "C:\ruta\documento.docx"
```

### Modo simulación (dry-run)

```batch
firmapdfs process --dry-run
```

### Solo procesar Word (sin PDF/firma)

```batch
firmapdfs process --no-pdf
```

### Procesar Word + PDF (sin firma)

```batch
firmapdfs process --no-sign
```

### Modo vigilancia (watch)

```batch
firmapdfs watch
```

### Ver estado del sistema

```batch
firmapdfs status
```

### Ver reporte

```batch
firmapdfs report
```

## 📊 Reporte CSV

El sistema genera un reporte CSV en `logs/report.csv` con:

| timestamp | filename | status | message | duration_ms |
|-----------|----------|--------|---------|-------------|
| 2024-01-15T10:30:00 | doc1.docx | OK | Procesado correctamente | 1234 |
| 2024-01-15T10:30:05 | doc2.docx | ERROR | Certificado no encontrado | 50 |
| 2024-01-15T10:30:10 | doc3.docx | SKIPPED | Ya procesado | 10 |

## 🔧 Comandos CLI

```
firmapdfs --help              # Ayuda general
firmapdfs process --help      # Ayuda del comando process
firmapdfs init                # Inicializar estructura
firmapdfs check               # Verificar configuración
firmapdfs check --fix         # Verificar y crear directorios
firmapdfs set-password        # Configurar contraseña
firmapdfs status              # Estado del sistema
firmapdfs report              # Ver reporte
```

## 🛡️ Seguridad

- ✅ **Contraseñas nunca en logs** - Filtro automático de datos sensibles
- ✅ **Windows Credential Manager** - Almacenamiento seguro de contraseñas
- ✅ **DPAPI** - Alternativa de cifrado de Windows
- ✅ **Validación de entrada** - Sanitización de rutas y nombres
- ✅ **Sin dependencias cloud** - Todo local

## 🐛 Solución de Problemas

### "Python no encontrado"

Asegúrate de que Python está en el PATH:
```batch
python --version
```

Si no funciona, reinstala Python marcando "Add to PATH".

### "Word COM error"

1. Verifica que Microsoft Word está instalado
2. Ejecuta como Administrador
3. Verifica que no hay instancias de Word abiertas

### "Certificado no encontrado"

1. Copia tu archivo .pfx a `config/certificate.pfx`
2. O actualiza la ruta en `config/config.yaml`

### "Error de firma"

1. Verifica la contraseña del certificado
2. Comprueba que el certificado no ha expirado
3. Ejecuta: `firmapdfs check`

## 📦 Desarrollo

### Ejecutar tests

```batch
pytest
```

### Formatear código

```batch
black src/
ruff check src/
```

### Verificar tipos

```batch
mypy src/
```

## 📄 Licencia

MIT License - Ver [LICENSE](LICENSE)

## 🗺️ Roadmap

- [x] **Sprint 0**: Setup y esqueleto del proyecto
- [ ] **Sprint 1**: Relleno Word mínimo viable
- [ ] **Sprint 2**: Export a PDF
- [ ] **Sprint 3**: Firma PDF mínima viable
- [ ] **Sprint 4**: Reglas configurables y robustez
- [ ] **Sprint 5**: Watcher + task scheduler
- [ ] **Sprint 6**: Hardening + DX

## 🤝 Contribuir

1. Fork del repositorio
2. Crear rama feature (`git checkout -b feature/nueva-funcion`)
3. Commit cambios (`git commit -am 'Añadir nueva función'`)
4. Push a la rama (`git push origin feature/nueva-funcion`)
5. Crear Pull Request
