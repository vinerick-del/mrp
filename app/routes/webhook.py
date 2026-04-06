"""
Rotas de Webhook - Recebimento de Sinais de Trading
"""

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from config import WEBHOOK_SECRET
from services.signal_service import process_signal
from services.telegram_service import send_telegram_message

webhook_router = APIRouter(prefix="/webhook", tags=["webhook"])


# Modelo de entrada
class SignalPayload(BaseModel):
    secret: str
    par: str
    tipo: str
    timeframe: str
    preco: float


@webhook_router.post("")
async def receive_signal(payload: SignalPayload):
    """
    Receber sinal de trading do TradingView

    - Valida WEBHOOK_SECRET
    - Processa o sinal com timestamp
    - Envia notificação via Telegram
    - Retorna confirmação
    """

    # Validar secret
    if payload.secret != WEBHOOK_SECRET:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Secret inválido"
        )

    # Processar sinal
    signal = process_signal(
        par=payload.par,
        tipo=payload.tipo,
        timeframe=payload.timeframe,
        preco=payload.preco
    )

    # Montar mensagem formatada
    message = f"""🚀 <b>{signal['tipo']}</b>
Par: <b>{signal['par']}</b>
Timeframe: <b>{signal['timeframe']}</b>
Preço: <b>{signal['preco']}</b>
Horário: <b>{signal['timestamp']}</b>"""

    # Enviar via Telegram
    telegram_result = send_telegram_message(message)

    # Retornar confirmação
    return {
        "status": "sinal recebido",
        "signal": signal,
        "telegram": telegram_result
    }
