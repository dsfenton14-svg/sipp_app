# Publicar actualizaciones con GitHub

1. Crear un repositorio privado en GitHub.
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
