"""
Rotas de Webhook - Recebimento de Sinais de Trading
"""

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from config import WEBHOOK_SECRET
from services.signal_service import process_signal
from services.telegram_service import send_telegram_message
from services.risk_management import validate_risk

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
    - Valida risco antes de enviar
    - Se aprovado: envia para Telegram
    - Se bloqueado: não envia e retorna motivo
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

    # Validar risco
    risk_validation = validate_risk(signal)

    # Se bloqueado por risco
    if not risk_validation["approved"]:
        print(f"⚠️ Sinal bloqueado: {risk_validation['reason']}")
        return {
            "status": "blocked",
            "reason": risk_validation["reason"],
            "signal": signal
        }

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
        "status": "sent",
        "reason": risk_validation["reason"],
        "signal": signal,
        "telegram": telegram_result
    }
