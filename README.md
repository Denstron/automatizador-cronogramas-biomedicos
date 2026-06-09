# 🏥 Automatizador de Cronogramas Biomédicos

Sistema web para gestión de mantenimiento de equipos médicos. Detecta vencimientos, clasifica por urgencia, genera informes y envía notificaciones automáticas.

## 🌐 Demo en vivo
👉 **[Probar la app aquí](https://biomedico-demo.streamlit.app/)**

---

## 📸 Capturas

![Login](<capturas de pantalla-inicio de sesión.png>)

![Dashboard](<capturas de pantalla-panel de control.png>)

---

## ✅ ¿Qué hace?

- Lee cronogramas de mantenimiento desde Excel
- Analiza más de 1000 equipos en segundos
- Clasifica cada equipo por estado:
  - 🔴 Mantenimiento este mes
  - 🟡 Próximo mes
  - 🟠 En 2 meses
  - 🟢 Sin urgencia
- Dashboard web con login privado por roles
- Técnicos marcan equipos como completados en tiempo real
- Jefe ve el progreso desde cualquier dispositivo
- Google Sheets como base de datos en tiempo real
- Genera reportes Excel con colores por área
- Envía emails automáticos con equipos del mes

---

## 🛠️ Tecnologías

- Python 3
- Streamlit
- pandas
- gspread + Google Sheets API
- OpenPyXL
- smtplib

---

## 📦 Versiones

| Versión | Descripción |
|---------|-------------|
| V1 | Análisis y reporte básico |
| V2 | Colores automáticos en Excel |
| V3 | Filtros por área + PDF |
| V4 | Notificaciones automáticas por email |
| V5 | App web con dashboard interactivo |
| V5.1 | Login privado + Google Sheets en tiempo real |

---

## 🏨 Caso de uso real

Desarrollado para **ESE Hospital San Rafael Yolombó**.
- 1056 equipos biomédicos
- 26 áreas del hospital
- Usado actualmente por el equipo de ingeniería biomédica

---

## 👤 Autor

**Daniel España Vargas** — Ingeniero Biomédico  
Registro INVIMA | Resolución 3100  
[github.com/Denstron](https://github.com/Denstron)
