# claude-windows

Programas de Windows (.exe) compilados automáticamente con GitHub Actions.

## Facturas

Programa nativo de Windows para llevar facturas: registro global con totales (facturado, cobrado, pendiente),
editor de factura, estados Pendiente/Pagada, exportar a CSV (Excel), respaldo e importación,
imprimir o guardar en PDF, y copiar la factura como imagen para enviarla por WhatsApp o correo.

- Código: carpeta `Facturas/` (C#, Windows Forms, .NET Framework 4.8, que ya viene con Windows 10/11).
- Datos: se guardan en `Documentos\Facturas\facturas.json`, con copias diarias en la subcarpeta `respaldo`.
  Se puede cambiar la carpeta desde el programa.
- El formato es el mismo de la versión web, así que los respaldos de una sirven en la otra.
- Descarga: pestaña **Releases** de este repositorio (`Facturas.exe`).

Cada vez que se sube un cambio a `main`, GitHub Actions compila una versión nueva y la publica en Releases.
