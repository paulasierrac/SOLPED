from Config.init_config import in_config
from repositories.TicketInsumo import TicketInsumoRepository
import socket

class TicketInsumoService:

    MAX_REINTENTOS = in_config("REINTENTOS") or 3

    def __init__(self, maquina: str, codigo: str):
        self.maquina = maquina or socket.gethostbyname()
        self.codigo = codigo

    def iniciar(self):

        ticket = TicketInsumoRepository.obtener_por_codigo(self.codigo)
        if not ticket:
            TicketInsumoRepository.crear(self.codigo, self.maquina)

        TicketInsumoRepository.actualizar_estado(
            self.codigo,
            estado="EN_PROCESO",
            observaciones="Inicio del procesamiento"
        )

    def finalizar(self):
        TicketInsumoRepository.actualizar_estado(
            self.codigo,
            estado="FINALIZADO",
            observaciones="Proceso finalizado correctamente",
            finalizar=True
        )

    def error(self, mensaje_error: str):
        ticket = TicketInsumoRepository.obtener_por_codigo(self.codigo)

        if not ticket:
            raise ValueError("Ticket no encontrado para manejar error")

        reintentos = ticket["numeroreintentos"] + 1
        estado = "REINTENTO" if reintentos < self.MAX_REINTENTOS else "ERROR"

        TicketInsumoRepository.actualizar_estado(
            self.codigo,
            estado=estado,
            observaciones=mensaje_error,
            incrementar_reintento=True
        )
