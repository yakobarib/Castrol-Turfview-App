# Castrol Turfview App

Aplicación web para procesar y transformar archivos de ventas y stock de Castrol al formato requerido por Turfview.

## Estructura del proyecto

```
1-Requerimientos/     → PDFs con especificaciones de Castrol
2-Prompts/            → Prompt de desarrollo y notas adicionales
3-Bases de Datos/     → Bases de datos de SKU, precios y clientes
4-Branding/           → Logos y paleta de colores
5-Ejemplos/           → Archivos de entrada y salida de ejemplo
6-Prueba/             → Archivos utilizados en pruebas
7-APP TURFVIEW/       → La aplicación (archivos HTML)
```

## Cómo usar la app

Abrir el archivo `7-APP TURFVIEW/turfview_castrol_app_v2_3.html` en el navegador y seguir las instrucciones en pantalla.

## ⚠️ Solución de errores por datos desactualizados

**Si la app da errores o resultados incorrectos**, la causa más probable es que las bases de datos estén desactualizadas.

Para solucionarlo, actualiza los siguientes archivos en la carpeta `3-Bases de Datos/` arrastrando las versiones nuevas:

- `BD_SKU_Codes.xlsx` — Códigos de producto
- `BD_Precios_Tarifa.xlsx` — Precios y tarifas
- `BD_CIFs_Clientes.xlsx` — Datos de clientes *(no incluido en el repositorio por contener datos sensibles)*

Una vez actualizados los archivos, vuelve a cargar la app en el navegador.

## Notas

- `BD_CIFs_Clientes.xlsx` no está incluido en este repositorio por contener datos fiscales de clientes. Debe mantenerse localmente y actualizarse de forma manual.
