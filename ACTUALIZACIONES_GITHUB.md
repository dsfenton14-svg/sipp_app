# Publicar actualizaciones con GitHub

1. Crear un repositorio privado en GitHub.
2. Subir el proyecto, excluyendo `.venv`, `.env`, archivos `.bin`, `respaldos` y `uploads`.
3. Revisar que no haya claves, contrasenas ni URLs privadas en el commit.
4. Crear una etiqueta para cada version estable:

   ```powershell
   git tag v2.3.1
   git push origin v2.3.1
   ```

5. GitHub Actions compilara `SiPP.exe` y creara el Release automaticamente.
6. Probar el instalador del Release antes de distribuirlo.

La aplicacion no debe descargar commits directamente. Las actualizaciones deben distribuirse como instaladores publicados en Releases. La URL del manifiesto se configurara cuando exista el repositorio remoto definitivo.
