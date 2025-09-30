import sys
import os
import time
import subprocess
import shutil
import zipfile
import psutil

KEEP_NAMES = {"venv", ".venv", "updater.py", "update_temp"}  # lo que NO se borra

def _kill_wait_parent(pid: int, timeout=12):
    try:
        if psutil.pid_exists(pid):
            p = psutil.Process(pid)
            p.wait(timeout=timeout)
    except psutil.NoSuchProcess:
        pass
    except Exception as e:
        print(f"No se pudo confirmar el cierre del proceso principal: {e}. Continuando de todas formas.")

def _find_zip_in_update_temp(update_dir: str) -> str | None:
    # Busca cualquier .zip en update_temp
    for name in os.listdir(update_dir):
        if name.lower().endswith(".zip"):
            return os.path.join(update_dir, name)
    return None

def _extract_zip(zip_path: str, extract_to: str) -> str:
    """Extrae el ZIP a extract_to y retorna el directorio raíz con contenido."""
    if os.path.exists(extract_to):
        shutil.rmtree(extract_to, ignore_errors=True)
    os.makedirs(extract_to, exist_ok=True)
    print(f"Extrayendo ZIP: {zip_path}")
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(extract_to)

    # Detección: zipball de GitHub suele tener un directorio raíz único
    entries = os.listdir(extract_to)
    if len(entries) == 1 and os.path.isdir(os.path.join(extract_to, entries[0])):
        return os.path.join(extract_to, entries[0])
    return extract_to

def _remove_all_except(target_dir: str, keep: set[str]):
    print("Eliminando archivos antiguos (excepto venv/.venv/updater.py/update_temp)...")
    for entry in os.listdir(target_dir):
        if entry in keep:
            continue
        full = os.path.join(target_dir, entry)
        try:
            if os.path.isdir(full):
                shutil.rmtree(full, ignore_errors=True)
            else:
                try:
                    os.remove(full)
                except PermissionError:
                    # Si está bloqueado, renombramos y borramos luego
                    tmp = full + ".old_del"
                    try:
                        os.replace(full, tmp)
                        os.remove(tmp)
                    except Exception:
                        pass
        except Exception as e:
            print(f"  -> Error al borrar '{entry}': {e}")

def _copy_all(src_dir: str, dst_dir: str):
    print("Copiando archivos nuevos...")
    for root, dirs, files in os.walk(src_dir):
        rel = os.path.relpath(root, src_dir)
        dest_root = os.path.join(dst_dir, rel) if rel != "." else dst_dir
        os.makedirs(dest_root, exist_ok=True)
        for d in dirs:
            os.makedirs(os.path.join(dest_root, d), exist_ok=True)
        for f in files:
            src_file = os.path.join(root, f)
            dst_file = os.path.join(dest_root, f)
            try:
                # Si el destino existe, lo reemplazamos
                if os.path.exists(dst_file):
                    try:
                        os.remove(dst_file)
                    except PermissionError:
                        tmp = dst_file + ".old_del"
                        try:
                            os.replace(dst_file, tmp)
                            os.remove(tmp)
                        except Exception:
                            pass
                shutil.copy2(src_file, dst_file)
            except Exception as e:
                print(f"  -> Error copiando '{src_file}' -> '{dst_file}': {e}")

def main():
    """
    Actualizador ZIP: espera al proceso principal, extrae update_temp/*.zip,
    elimina TODO excepto venv/.venv/updater.py/update_temp y copia los archivos nuevos.
    Luego actualiza version.txt, limpia temporales y relanza install.bat.
    """
    if len(sys.argv) < 3:
        print("Error: No se proporcionó el PID del proceso principal y la nueva versión.")
        time.sleep(5)
        return

    pid = int(sys.argv[1])
    new_version = sys.argv[2]
    target_dir = os.getcwd()
    update_dir = os.path.join(target_dir, "update_temp")
    extracted_dir = os.path.join(update_dir, "_extracted")

    print(f"Actualizador iniciado. Nueva versión a instalar: {new_version}")
    print(f"Esperando a que el proceso principal (PID: {pid}) se cierre...")
    _kill_wait_parent(pid, timeout=15)

    time.sleep(1.5)

    if not os.path.isdir(update_dir):
        print(f"Error: El directorio de actualización '{update_dir}' no existe.")
        time.sleep(5)
        return

    zip_path = _find_zip_in_update_temp(update_dir)
    if not zip_path:
        print("Error: No se encontró ningún archivo .zip dentro de update_temp.")
        time.sleep(5)
        return

    try:
        src_root = _extract_zip(zip_path, extracted_dir)
    except Exception as e:
        print(f"No fue posible extraer el ZIP: {e}")
        time.sleep(5)
        return

    # Borrar todo excepto venv/.venv/updater.py/update_temp
    _remove_all_except(target_dir, KEEP_NAMES)

    # Copiar nuevos archivos
    _copy_all(src_root, target_dir)

    # Escribir versión
    try:
        with open(os.path.join(target_dir, "version.txt"), "w", encoding="utf-8") as f:
            f.write(new_version)
        print(f"Archivo de versión actualizado a '{new_version}'.")
    except Exception as e:
        print(f"No se pudo escribir version.txt: {e}")

    # Limpieza
    try:
        shutil.rmtree(extracted_dir, ignore_errors=True)
    except Exception:
        pass
    try:
        # Mantener update_temp por si se quiere auditar el zip descargado
        pass
    except Exception:
        pass

    # Relanzar la app (instalador por si hay requirements)
    try:
        print("Reiniciando la aplicación...")
        subprocess.Popen(["install.bat"])
    except Exception as e:
        print(f"No se pudo relanzar install.bat: {e}")

    print("Actualización completada. El actualizador se cerrará ahora.")

if __name__ == '__main__':
    main()
