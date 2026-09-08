# Publicar actualizaciones con GitHub

1. Mantener un repositorio publico en GitHub para que SiPP pueda consultar los Releases sin credenciales.
2. Subir el proyecto, excluyendo `.venv`, `.env`, archivos `.bin`, `respaldos` y `uploads`.
3. Revisar que no haya claves, contrasenas ni URLs privadas en el commit.
4. Crear una etiqueta para cada version estable:

   ```powershell
   git tag v2.3.1
   git push origin v2.3.1
   ```

5. GitHub Actions compilara `SiPP.exe`, lo empaquetara con Inno Setup y creara el instalador `SiPP-Setup-vX.Y.Z.exe` en el Release automaticamente.
6. Probar el instalador del Release antes de distribuirlo.

La aplicacion no debe descargar commits directamente. Las actualizaciones deben distribuirse como instaladores publicados en Releases. El instalador conserva la configuracion local del usuario al actualizar y crea accesos directos y desinstalador de Windows.

## Distribucion gratuita sin firma digital

- Usar solamente el repositorio oficial: https://github.com/dsfenton14-svg/sipp_app
- Descargar el instalador desde GitHub Releases, no desde archivos adjuntos ni enlaces alternativos.
- Verificar `SHA256SUMS.txt` antes de instalar.
- Enviar cada instalador a Microsoft Defender para analisis: https://www.microsoft.com/en-us/wdsi/filesubmission
- Si SmartScreen bloquea una descarga conocida, el usuario puede abrir Propiedades, marcar Desbloquear y ejecutar como administrador.
- No desactivar Defender o SmartScreen globalmente.

La firma digital no esta incluida en este flujo gratuito. Por eso Windows puede mostrar una advertencia en cada instalador nuevo aunque el archivo sea legitimo.
