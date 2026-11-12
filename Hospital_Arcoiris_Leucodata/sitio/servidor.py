import html
import secrets
import time

from http.cookies import SimpleCookie
from http.server import HTTPServer
from urllib.parse import parse_qs, urlsplit

import app
import captura

from autenticacion import comprobar_credenciales, leer_tabla


# Las sesiones viven en memoria.
# Al cerrar Python, todos deberán iniciar sesión nuevamente.
SESIONES = {}

# Intentos fallidos de acceso por dirección.
INTENTOS = {}

DURACION_SESION = 30 * 60  # 30 minutos

ROLES_VALIDOS = {
    "Administrador",
    "Medico",
    "Personal",
    "Paciente",
    "Tutor",
}


class Servidor(app.Handler):

    def redirigir(self, destino, cookie=None):
        """Indica al navegador qué pantalla debe abrir."""

        self.send_response(303)
        self.send_header("Location", destino)
        self.send_header("Cache-Control", "no-store")

        if cookie:
            self.send_header("Set-Cookie", cookie)

        self.send_header("Content-Length", "0")
        self.end_headers()

    def obtener_token(self):
        """Lee el identificador de sesión del navegador."""

        cookie = SimpleCookie()

        try:
            cookie.load(self.headers.get("Cookie", ""))
            valor = cookie.get("sesion")
            return valor.value if valor else None
        except Exception:
            return None

    def cuenta_actual(self):
        """Comprueba la sesión y el estado actual de la cuenta."""

        ahora = time.time()

        # Elimina sesiones vencidas.
        for token, sesion in list(SESIONES.items()):
            if sesion["vence"] <= ahora:
                SESIONES.pop(token, None)

        token = self.obtener_token()
        sesion = SESIONES.get(token)

        if not sesion:
            return None

        libro = app.load()

        try:
            cuentas = (
                leer_tabla(libro, "Cuentas_Personal")
                + leer_tabla(libro, "Cuentas_Pacientes")
            )

            coincidencias = [
                cuenta
                for cuenta in cuentas
                if cuenta.get("cuenta_id") == sesion["cuenta_id"]
            ]

            if len(coincidencias) == 1:
                cuenta = coincidencias[0]

                if (
                    cuenta.get("estado") == "Activo"
                    and cuenta.get("rol") in ROLES_VALIDOS
                ):
                    return cuenta

        finally:
            libro.close()

        SESIONES.pop(token, None)
        return None

    def do_GET(self):
        if not self.local_host():
            return self.json(403, {
                "error": "Abre el sitio desde su dirección local."
            })

        ruta = urlsplit(self.path).path

        try:
            # La captura general permite modificar información:
            # únicamente una cuenta administradora puede utilizarla.
            if ruta in ("/", "/api/form"):
                cuenta = self.cuenta_actual()

                if not cuenta:
                    return self.redirigir("/hospital/")

                if cuenta.get("rol") != "Administrador":
                    return self.json(403, {
                        "error": "Esta sección requiere administración."
                    })

            if ruta == "/panel/":
                cuenta = self.cuenta_actual()

                if not cuenta:
                    return self.redirigir("/hospital/")

                usuario = html.escape(str(cuenta["usuario"]))
                rol = html.escape(str(cuenta["rol"]))

                pagina = f"""
                <!DOCTYPE html>
                <html lang="es">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport"
                          content="width=device-width, initial-scale=1">
                    <title>Mi espacio | Hospital Arcoíris</title>
                    <link rel="stylesheet"
                          href="/hospital/estilos.css">
                </head>
                <body>
                    <main class="contenedor">
                        <section class="presentacion">
                            <img
                                src="/hospital/Hospital.jpeg"
                                alt="Hospital Arcoíris"
                                class="logo-hospital"
                            >

                            <h1>Bienvenido</h1>
                            <p>Usuario: {usuario}</p>
                            <p>Rol: {rol}</p>
                        </section>

                        <section class="tarjeta">
                            <h2>Sesión iniciada</h2>

                            <p>
                                Tu cuenta fue comprobada.
                                Aquí agregaremos los módulos
                                correspondientes a tus permisos.
                            </p>

                            <form action="/salir" method="post">
                                <button type="submit">
                                    Cerrar sesión
                                </button>
                            </form>
                        </section>
                    </main>
                </body>
                </html>
                """

                return self.reply(
                    200,
                    pagina.encode("utf-8"),
                    "text/html; charset=utf-8",
                )

            # Conserva las páginas del hospital y del registro.
            return super().do_GET()

        except (OSError, ValueError, KeyError):
            return self.json(503, {
                "error": "No se pudo consultar el Excel."
            })

    def do_POST(self):
        # Comprueba que el formulario venga de nuestro sitio.
        if (
            not self.local_host()
            or self.headers.get("Origin") != captura.URL
        ):
            return self.json(403, {
                "error": "Abre el formulario desde el servidor local."
            })

        ruta = urlsplit(self.path).path

        try:
            if ruta == "/salir":
                token = self.obtener_token()
                SESIONES.pop(token, None)

                return self.redirigir(
                    "/hospital/",
                    "sesion=; Path=/; HttpOnly; "
                    "SameSite=Strict; Max-Age=0",
                )

            if ruta == "/login":
                direccion = self.client_address[0]
                ahora = time.time()

                fallos = [
                    instante
                    for instante in INTENTOS.get(direccion, [])
                    if ahora - instante < 300
                ]

                if len(fallos) >= 10:
                    return self.json(429, {
                        "error": "Espera cinco minutos antes de reintentar."
                    })

                if (
                    self.headers.get_content_type()
                    != "application/x-www-form-urlencoded"
                ):
                    return self.json(415, {
                        "error": "Formato de formulario incorrecto."
                    })

                longitud = int(
                    self.headers.get("Content-Length", "0")
                )

                if not 0 < longitud <= 10000:
                    return self.json(400, {
                        "error": "Tamaño de formulario incorrecto."
                    })

                contenido = self.rfile.read(longitud).decode("utf-8")
                datos = parse_qs(contenido)

                usuario = datos.get("usuario", [""])[0].strip()
                contrasena = datos.get("contrasena", [""])[0]

                cuenta = comprobar_credenciales(
                    usuario,
                    contrasena,
                )

                if not cuenta:
                    INTENTOS[direccion] = fallos + [ahora]

                    return self.json(401, {
                        "error": "Usuario o contraseña incorrectos. "
                                 "Vuelve atrás para intentarlo de nuevo."
                    })

                if cuenta.get("estado") != "Activo":
                    return self.json(403, {
                        "error": "La cuenta todavía no está habilitada."
                    })

                if cuenta.get("rol") not in ROLES_VALIDOS:
                    return self.json(403, {
                        "error": "La cuenta no tiene un rol reconocido."
                    })

                INTENTOS.pop(direccion, None)

                # Sustituye cualquier sesión anterior del navegador.
                SESIONES.pop(self.obtener_token(), None)

                token = secrets.token_urlsafe(32)

                SESIONES[token] = {
                    "cuenta_id": cuenta["cuenta_id"],
                    "vence": ahora + DURACION_SESION,
                }

                return self.redirigir(
                    "/panel/",
                    f"sesion={token}; Path=/; HttpOnly; "
                    f"SameSite=Strict; Max-Age={DURACION_SESION}",
                )

            if ruta == "/api/form":
                cuenta = self.cuenta_actual()

                if not cuenta or cuenta.get("rol") != "Administrador":
                    return self.json(403, {
                        "error": "Esta acción requiere administración."
                    })

            # Conserva el guardado de registros de tu programa.
            return super().do_POST()

        except (ValueError, UnicodeError):
            return self.json(400, {
                "error": "Revisa los datos enviados."
            })

        except (OSError, KeyError):
            return self.json(503, {
                "error": "No se pudo consultar el Excel."
            })


if __name__ == "__main__":
    with HTTPServer(("127.0.0.1", 8765), Servidor) as servidor:
        print("Hospital: http://127.0.0.1:8765/hospital/")
        print("Registro: http://127.0.0.1:8765/registro/")
        print("Presiona Ctrl+C para detener el servidor.")

        try:
            servidor.serve_forever()
        except KeyboardInterrupt:
            pass
        