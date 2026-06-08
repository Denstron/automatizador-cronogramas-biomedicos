# 🏥 Automatizador de Cronogramas Biomédicos

Sistema automático para gestión de mantenimiento de equipos médicos.
Detecta vencimientos, clasifica por urgencia, genera reportes y envía 
notificaciones automáticas por email.

## ¿Qué hace?

- Lee cronogramas de mantenimiento desde Excel
- Analiza 1000+ equipos en segundos
- Clasifica cada equipo por estado:
  - 🔴 Mantenimiento este mes
  - 🟡 Mantenimiento próximo mes
  - 🟠 Mantenimiento en 2 meses
  - 🟢 Sin urgencia
- Genera reporte Excel con colores y 26 hojas por área
- Genera reporte PDF profesional
- Envía email automático con equipos del mes y archivos adjuntos
- Soporta múltiples destinatarios (técnicos por área)

## Tecnologías

- Python 3
- pandas
- openpyxl
- reportlab
- smtplib

## Versiones

- V1: Análisis y reporte básico
- V2: Colores automáticos en Excel
- V3: Filtros por área + reporte PDF
- V4: Notificaciones automáticas por email

## Caso de uso real

Desarrollado para ESE Hospital San Rafael Yolombó.
Procesando 1056 equipos biomédicos en 26 áreas.
Usado actualmente por el equipo de ingeniería biomédica.

## Autor

Daniel España Vargas — Ingeniero Biomédico  
Registro INVIMA | Resolución 3100  
github.com/Denstron
