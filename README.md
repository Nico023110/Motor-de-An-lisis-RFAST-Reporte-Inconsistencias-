# Motor de Análisis RFAST - Reporte de Inconsistencias

Sistema automatizado de auditoría y validación de datos en salud (RIPS / Listado de actividades de usuarios) para la detección de inconsistencias en codificación de causas externas, finalidades RIPS y tipos de consulta por centro de producción.

## 📁 Estructura del Proyecto

```
reporte_inconsistencias/
├── fuente/                      # Archivos de entrada (.xlsx)
├── reporte/                     # Informes de inconsistencias generados (.xlsx)
├── reporte_incompletos/         # Registros incompletos / pendientes
└── scripts/
    └── motor_reporte V1.py      # Script principal del motor de validación
```

## 🚀 Requisitos e Instalación

### Requisitos Previos
* Python 3.8+

### Librerías Necesarias
```bash
pip install pandas openpyxl
```

## ⚙️ Uso

1. Coloca el archivo de datos Excel de entrada dentro de la carpeta `fuente/`.
2. Ejecuta el script principal:
   ```bash
   python "scripts/motor_reporte V1.py"
   ```
3. El reporte generado con estilos e inconsistencias resaltadas en rojo se guardará automáticamente en la carpeta `reporte/`.

## 🔍 Validaciones Incorporadas

1. **Validación 01**: Causa Externa Incorrecta PyM (Promoción y Mantenimiento).
2. **Validación 02**: Finalidad Incorrecta en Programas de Control (Hipertensión/Diabetes).
3. **Validación 03**: Finalidad Incorrecta en Consultas de Primera Vez.
4. **Validación 04**: Finalidad Incorrecta en Controles de Odontología.
5. **Validación 05**: Finalidad Incorrecta en Planificación Familiar.
6. **Validación 06**: Finalidad Incorrecta en Programas de Detección Temprana.
7. **Validación 07**: Finalidad Incorrecta en Educación Individual.
