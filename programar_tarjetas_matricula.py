#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import time
import threading
from smartcard.System import readers
from smartcard.CardMonitoring import CardMonitor, CardObserver
from smartcard.Exceptions import NoCardException

class NdefManager:
    """
    Este programa está pensado para tarjetas ISO 14443-3A NXP-NTAG213 Type A de 180 bytes,
    con hasta 130 caracteres en la URL a escribir.
    """
    def __init__(self):
        """Constructor. Se conecta al primer lector encontrado."""
        r = readers()
        if len(r) == 0:
            print("No se encontraron lectores")
            sys.exit()
        print("Lectores disponibles:")
        for i, reader in enumerate(r):
            print("  {}: {}".format(i, reader))
        self.lector = r[0]
        print("Se selecciona el lector:", self.lector)
        self.conexion = None  # Se asignará al conectar la tarjeta

    def esperar_tarjeta(self):
        """Espera a que se inserte la tarjeta y establece la conexión."""
        cardInsertedEvent = threading.Event()

        class WaitCardObserver(CardObserver):
            def update(self, observable, cards):
                addedCards, removedCards = cards
                if addedCards:
                    cardInsertedEvent.set()

        cardMonitor = CardMonitor()
        observer = WaitCardObserver()
        cardMonitor.addObserver(observer)
        print("Esperando a que se coloque la tarjeta...")
        cardInsertedEvent.wait()
        cardMonitor.deleteObserver(observer)
        conexion = self.lector.createConnection()
        try:
            conexion.connect()
        except NoCardException as e:
            print("Error al conectar con la tarjeta:", e)
            return None
        print("Tarjeta conectada en lector:", self.lector)
        return conexion

    def esperar_remocion(self):
        """Espera a que se retire la tarjeta."""
        cardRemovedEvent = threading.Event()

        class WaitCardRemovalObserver(CardObserver):
            def update(self, observable, cards):
                addedCards, removedCards = cards
                if removedCards:
                    cardRemovedEvent.set()

        cardMonitor = CardMonitor()
        observer = WaitCardRemovalObserver()
        cardMonitor.addObserver(observer)
        print("Esperando a que se retire la tarjeta...")
        cardRemovedEvent.wait()
        cardMonitor.deleteObserver(observer)
        print("Tarjeta retirada.")

    def _escribir_pagina(self, pagina, datos):
        """
        Escribe 4 bytes (lista de 4 enteros) en la página indicada.
        Se usa el comando:
          FF 00 00 00 06 D4 42 A2 [pagina] [dato0] [dato1] [dato2] [dato3]
        """
        if len(datos) != 4:
            raise ValueError("Los datos deben ser 4 bytes")
        apdu = [0xFF, 0x00, 0x00, 0x00, 0x06,
                0xD4, 0x42, 0xA2, pagina] + datos
        respuesta, sw1, sw2 = self.conexion.transmit(apdu)
        if sw1 == 0x90 and sw2 == 0x00:
            pass
        else:
            print("Error al escribir la página {}: SW1={} SW2={}".format(pagina, hex(sw1), hex(sw2)))
        return respuesta, sw1, sw2

    def _leer_bloque(self, pagina_inicio):
        """
        Lee 4 páginas (16 bytes de datos reales) a partir de la página indicada.
        Se utiliza el comando:
          FF 00 00 00 04 D4 42 30 [pagina_inicio]
        Como la respuesta incluye un encabezado de longitud variable, extraemos los 16 bytes finales.
        """
        apdu = [0xFF, 0x00, 0x00, 0x00, 0x04,
                0xD4, 0x42, 0x30, pagina_inicio]
        respuesta, sw1, sw2 = self.conexion.transmit(apdu)
        if sw1 == 0x90 and sw2 == 0x00:
            if len(respuesta) > 16:
                return respuesta[-16:]
            else:
                return respuesta
        else:
            print("Error al leer bloque a partir de la página {}: SW1={} SW2={}".format(pagina_inicio, hex(sw1), hex(sw2)))
            return None

    def escribir_ndef(self, url):
        """
        Crea y escribe un mensaje NDEF en el chip usando la URL dada.
        La estructura es:
          [0x03, LEN, 0xD1, 0x01, PAYLOAD_LEN, 0x55, 0x04, ... URL en ASCII ..., 0xFE]
        Se escribe a partir de la página 4, en bloques de 4 bytes.
        """
        # Preparar la URL: quitar el esquema y usar 0x04 para "https://"
        prefijo = 0x04
        if url.startswith("https://"):
            url_sin_prefijo = url[len("https://"):]
        elif url.startswith("http://"):
            url_sin_prefijo = url[len("http://"):]
        else:
            url_sin_prefijo = url
        url_bytes = list(url_sin_prefijo.encode('utf-8'))
        payload = [prefijo] + url_bytes

        # Registro NDEF: [0xD1, 0x01, payload_len, 0x55, payload]
        registro_ndef = [0xD1, 0x01, len(payload), 0x55] + payload
        # TLV: [0x03, LEN, registro_ndef, 0xFE]
        tlv = [0x03, len(registro_ndef)] + registro_ndef + [0xFE]
        # Rellenar hasta que la longitud sea múltiplo de 4 (una página = 4 bytes)
        if len(tlv) % 4 != 0:
            relleno = 4 - (len(tlv) % 4)
            tlv += [0x00] * relleno

        # Dividir el TLV en bloques (cada bloque de 4 bytes)
        paginas = [tlv[i:i+4] for i in range(0, len(tlv), 4)]
        pagina_inicio = 4
        for i, datos in enumerate(paginas):
            self._escribir_pagina(pagina_inicio + i, datos)
        print("Escritura completada.")

    def leer_ndef(self):
        """
        Lee el mensaje NDEF de forma dinámica, recorriendo bloques (4 páginas = 16 bytes)
        desde la página 4 hasta encontrar el terminador 0xFE o alcanzar un máximo (página 40).
        Luego se busca el TLV 0x03 y se extrae el mensaje.
        Retorna la URL leída o None en caso de error.
        """
        time.sleep(0.5)  # Espera para que la escritura se asiente
        datos_leidos = []
        current_page = 4
        max_page = 40  # Ajustable según la capacidad del chip
        while current_page < max_page:
            bloque = self._leer_bloque(current_page)
            if bloque is None:
                break
            datos_leidos.extend(bloque)
            if 0xFE in bloque:
                break
            current_page += 4

        try:
            indice = datos_leidos.index(0x03)
        except ValueError:
            if datos_leidos and all(datos_leidos[i] == [0xD5, 0x43, 0x01][i % 3] for i in range(len(datos_leidos))):
                print("El chip parece estar defectuoso en la lectura.")
            else:
                print("No se encontró un TLV NDEF (0x03) en la memoria.")
            return None

        if indice + 1 >= len(datos_leidos):
            print("No se encontró longitud NDEF tras el tag 0x03.")
            return None

        ndef_len = datos_leidos[indice + 1]
        mensaje_ndef = datos_leidos[indice + 2:indice + 2 + ndef_len]

        if len(mensaje_ndef) >= 5:
            payload_len = mensaje_ndef[2]
            if len(mensaje_ndef) >= 4 + payload_len:
                payload = mensaje_ndef[4:4 + payload_len]
                codigo_prefijo = payload[0]
                mapa_prefijos = {
                    0x00: "",
                    0x01: "http://www.",
                    0x02: "https://www.",
                    0x03: "http://",
                    0x04: "https://",
                }
                prefijo_str = mapa_prefijos.get(codigo_prefijo, "")
                url_leida = prefijo_str + "".join(chr(b) for b in payload[1:])
                print("URL leída:", url_leida)
                return url_leida
            else:
                print("El payload no tiene la longitud esperada.")
                return None
        else:
            print("El mensaje NDEF es muy corto para analizarlo.")
            return None

    # Método 1: Escribir y leer una lista de matrículas.
    def escribir_y_leer_lista(self, lista_matriculas):
        for matricula in lista_matriculas:
            print("\n--- Proceso para matrícula:", matricula, "---")
            print("Por favor, coloque la tarjeta para escribir la matrícula.")
            conexion = self.esperar_tarjeta()
            if conexion is None:
                print("No se pudo conectar la tarjeta.")
                continue
            self.conexion = conexion
            print("Escribiendo matrícula:", matricula)
            self.escribir_ndef(matricula)
            print("Leyendo matrícula:")
            leida = self.leer_ndef()
            if leida is not None:
                print("Confirmación: matrícula leída:", leida)
            else:
                print("Error en la lectura.")
            self.esperar_remocion()
            print("Tarjeta procesada.")

    # Método 2: Escribir y leer una sola matrícula.
    def escribir_y_leer(self, matricula):
        print("\nPor favor, coloque la tarjeta para escribir la matrícula:", matricula)
        conexion = self.esperar_tarjeta()
        if conexion is None:
            print("No se pudo conectar la tarjeta.")
            return
        self.conexion = conexion
        print("Escribiendo matrícula:", matricula)
        self.escribir_ndef(matricula)
        print("Leyendo matrícula:")
        leida = self.leer_ndef()
        if leida is not None:
            print("Confirmación: matrícula leída:", leida)
        else:
            print("Error en la lectura.")
        self.esperar_remocion()

    # Método 3: Leer una sola matrícula.
    def leer_una(self):
        print("\nPor favor, coloque la tarjeta para leer la matrícula:")
        conexion = self.esperar_tarjeta()
        if conexion is None:
            print("No se pudo conectar la tarjeta.")
            return None
        self.conexion = conexion
        leida = self.leer_ndef()
        self.esperar_remocion()
        return leida

    # Método 4: Leer todas las tarjetas de una lista de matrículas esperadas, confirmando cada una.
    def leer_todas(self, lista_matriculas):
        for matricula in lista_matriculas:
            print("\n--- Confirmación para matrícula esperada:", matricula, "---")
            print("Por favor, coloque la tarjeta para confirmar la matrícula.")
            conexion = self.esperar_tarjeta()
            if conexion is None:
                print("No se pudo conectar la tarjeta.")
                continue
            self.conexion = conexion
            leida = self.leer_ndef()
            if leida is not None:
                if leida == f"https://{matricula}":
                    print("Confirmación: la matrícula coincide:", leida)
                else:
                    print("Error: se esperaba '{}' pero se leyó '{}'".format(matricula, leida))
            else:
                print("Error en la lectura.")
            self.esperar_remocion()
            print("Tarjeta procesada.")

if __name__ == '__main__':
    ndef_manager = NdefManager()
    lista_matriculas = [
    "24E0300585", "23E0300135", "24E0300711", "24E0300189", "24E0300333",
    "24E0300126", "24E0300747", "24E0300180", "23E0300765", "24E0300270",
    "24E0300414", "22E0300468", "24E0300675", "24E0300054", "24E0300108",
    "24E0300351", "24E0300207", "24E0300657", "24E0300567", "24E0300513",
    "24E0300027", "24E0300405", "23E0300756", "24E0300396", "24E0300486",
    "24E0300612", "24E0300018", "23E0300252", "24E0300342", "24E0300765",
    "22E0300306", "24E0300639", "24E0300432", "24E0300504", "24E0300603",
    "24E0300459", "24E0300549", "24E0300072", "23E0300774", "24E0300783",
    "24E0300792", "24E0300144", "22E0300153", "24E0300099", "24E0300522",
    "23E0300504", "24E0300450", "24E0300774", "23E0300666", "24E0300684",
    "24E0300576", "23E0300072", "21E0300666", "24E0300225", "24E0300387",
    "24E0300279", "24E0300315", "24E0300288", "24E0300693", "24E0300297",
    "24E0300477", "24E0300558", "24E0300729", "24E0300468", "23E0300036",
    "24E0300243", "24E0300801", "23E0300315", "24E0301422", "24E0300531",
    "24E0300378", "24E0300756", "24E0300135", "24E0300738", "24E0300369",
    "24E0300648", "24E0300666", "24E0300009", "22E0300090", "24E0300216",
    "24E0300702", "24E0300171", "24E0300594", "24E0300234", "22E0300252",
    "23E0300351", "24E0300081", "24E0300621", "24E0300720", "24E0300045",
    "24E0300324", "24E0300441", "24E0300162", "24E0300198", "24E0300540",
    "24E0300495", "24E0300306", "24E0300360", "23E0300153", "24E0300090",
    "23E0300432"
    ]
    # --- Selecciona descomentando el método que necesites ---
    matriculas_faltantes = ["24E0300360"]
    # Método 1: Escribir y leer una lista de matrículas.
    #lista_matriculas = ["24E0300585", "23E0300135", "24E0300711"]
    ndef_manager.escribir_y_leer_lista(matriculas_faltantes)

    # Método 2: Escribir y leer una sola matrícula.
    #matricula = "24E0300585"
    #ndef_manager.escribir_y_leer(matricula)

    # Método 3: Leer una sola matrícula.
    #leida = ndef_manager.leer_una()

    # Método 4: Leer todas las tarjetas de una lista de matrículas esperadas.
    #lista_matriculas = ["24E0300585", "23E0300135", "24E0300711"]
    #ndef_manager.leer_todas(lista_matriculas)
