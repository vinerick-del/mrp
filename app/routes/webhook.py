"""
Rotas de Webhook - Recebimento de Sinais de Trading
"""

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from config import WEBHOOK_SECRET

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
    - Se inválido retorna 401
    - Se válido retorna confirmação de recebimento
    """

    # Validar secret
    if payload.secret != WEBHOOK_SECRET:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Secret inválido"
        )

    # Sinal válido recebido
    return {
        "status": "sinal recebido",
        "par": payload.par,
        "tipo": payload.tipo,
        "timeframe": payload.timeframe,
        "preco": payload.preco
    }
