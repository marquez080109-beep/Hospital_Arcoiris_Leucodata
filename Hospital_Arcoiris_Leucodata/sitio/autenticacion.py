import hashlib
import hmac
import getpass

from openpyxl.utils.cell import range_boundaries # pyright: ignore[reportMissingModuleSource]

import app


def leer_tabla(libro, nombre_hoja):
    """Lee los registros de una tabla del Excel."""

    if nombre_hoja not in libro.sheetnames:
        return []

    hoja = libro[nombre_hoja]
    nombre_tabla = "tbl_" + nombre_hoja

    if nombre_tabla not in hoja.tables:
        return []

    tabla = hoja.tables[nombre_tabla]
    izquierda, arriba, derecha, abajo = range_boundaries(tabla.ref)

    encabezados = [
        hoja.cell(arriba, columna).value
        for columna in range(izquierda, derecha + 1)
    ]

    registros = []

    for fila in hoja.iter_rows(
        min_row=arriba + 1,
        max_row=abajo,
        min_col=izquierda,
        max_col=derecha,
        values_only=True,
    ):
        if fila[0] is not None:
            registros.append(dict(zip(encabezados, fila)))

    return registros


def comprobar_credenciales(usuario, contrasena):
    """Comprueba la contraseña sin revelar ni modificar su huella."""

    libro = app.load()

    try:
        credenciales = leer_tabla(libro, "Credenciales_Web")

        cuentas = (
            leer_tabla(libro, "Cuentas_Personal")
            + leer_tabla(libro, "Cuentas_Pacientes")
        )

        # Buscamos una única credencial para ese usuario.
        coincidencias = [
            registro
            for registro in credenciales
            if registro.get("usuario") == usuario
        ]

        if len(coincidencias) != 1:
            return None

        credencial = coincidencias[0]

        # Debe coincidir con el algoritmo de tu formulario de registro.
        if credencial.get("algoritmo") != "scrypt-131072-8-1":
            return None

        try:
            sal = bytes.fromhex(credencial["sal_hex"])
            huella_guardada = bytes.fromhex(credencial["huella_hex"])

            huella_recibida = hashlib.scrypt(
                contrasena.encode("utf-8"),
                salt=sal,
                n=2**17,
                r=8,
                p=1,
                maxmem=256 * 1024**2,
            )

        except (ValueError, TypeError, KeyError):
            return None

        if not hmac.compare_digest(
            huella_recibida,
            huella_guardada,
        ):
            return None

        # Relacionamos la contraseña con la cuenta correspondiente.
        cuentas_encontradas = [
            cuenta
            for cuenta in cuentas
            if cuenta.get("cuenta_id") == credencial.get("cuenta_id")
            and cuenta.get("usuario") == usuario
        ]

        if len(cuentas_encontradas) != 1:
            return None

        return cuentas_encontradas[0]

    finally:
        libro.close()


# Esta parte solo se ejecuta al probar el archivo desde la terminal.
if __name__ == "__main__":
    print("PRUEBA DE CREDENCIALES DEL HOSPITAL")

    usuario = input("Usuario: ").strip()
    contrasena = getpass.getpass("Contraseña: ")

    try:
        cuenta = comprobar_credenciales(usuario, contrasena)

        if cuenta is None:
            print("Usuario o contraseña incorrectos.")
        else:
            print("\nLa contraseña corresponde a esta cuenta.")
            print("Rol:", cuenta.get("rol"))
            print("Estado:", cuenta.get("estado"))
            print(
                "\nEsta prueba no inicia una sesión "
                "ni activa la cuenta."
            )

    except OSError:
        print("No se pudo abrir el Excel. Revisa su ubicación.")
