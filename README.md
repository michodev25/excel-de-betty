# Listado de precios

Aplicación Astro para actualizar los productos del listado mayorista y descargar un Excel con el estilo de la hoja original.

Requiere Node.js 22.19 o posterior (también funciona con Node.js 24).

## Iniciar

```bash
npm install
npm run dev
```

Abre la dirección local que muestra Astro. Para generar una versión estática: `npm run build`.

## Uso

- Busca por producto o proveedor y filtra por categoría o artículos sin precio.
- Cambia el precio directamente en la tabla o en la tarjeta móvil.
- Usa **Añadir producto** o **Editar datos** para modificar los demás campos.
- Elige la **Fecha del listado** antes de descargar. La fecha aparece en el nombre del archivo y la pestaña del Excel.
- Los cambios se guardan automáticamente en el almacenamiento de este navegador. Para continuar en otro dispositivo, descarga el Excel y usa **Importar Excel** allí. La importación sustituye el catálogo local después de confirmarlo.

El catálogo inicial procede de `Listado de precios 24-09-26.xlsx`. El archivo fuente no se modifica.
