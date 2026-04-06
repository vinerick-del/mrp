"""
Rotas de Webhook - Orquestrador do Sistema de Trading
Fluxo completo: Signal → Risk → MT5 → Telegram
"""

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from typing import Optional, List
from config import WEBHOOK_SECRET
from services.signal_service import process_signal
from services.strategy_service import generate_signal
from services.risk_management import validate_risk
from services.mt5_service import send_order, connect_mt5
from services.telegram_service import send_telegram_message

webhook_router = APIRouter(prefix="/webhook", tags=["webhook"])


# Modelos de entrada
class SignalPayload(BaseModel):
    secret: str
    par: str
    tipo: str
    timeframe: str
    preco: float


class MarketDataPayload(BaseModel):
    secret: str
    pair: str
    timeframe: str
    close_prices: List[float]


@webhook_router.post("")
async def receive_signal(payload: SignalPayload):
    """
    Orquestrador: Receber sinal direto e processar

    Fluxo: Secret → Signal → Risk → MT5 → Telegram
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

    # Conectar ao MT5 e executar ordem
    mt5_result = send_order({
        "pair": signal["par"],
        "signal": signal["tipo"]
    })

    # Se erro no MT5
    if mt5_result["status"] == "erro":
        print(f"❌ Erro MT5: {mt5_result['message']}")
        return {
            "status": "blocked",
            "reason": f"Erro ao executar ordem: {mt5_result['message']}",
            "signal": signal
        }

    # Montar mensagem Telegram
    message = f"""🚀 <b>{signal['tipo']}</b>
Par: <b>{signal['par']}</b>
Timeframe: <b>{signal['timeframe']}</b>
Preço: <b>{signal['preco']}</b>
SL: <b>{mt5_result.get('sl', 'N/A')}</b>
TP: <b>{mt5_result.get('tp', 'N/A')}</b>
Status: <b>EXECUTADO</b>"""

    # Enviar Telegram
    telegram_result = send_telegram_message(message)

    print(f"✅ Ordem executada: {signal['tipo']} {signal['par']}")

    return {
        "status": "executed",
        "reason": "Ordem executada com sucesso",
        "signal": signal,
        "order": mt5_result,
        "telegram": telegram_result
    }


@webhook_router.post("/market-data")
async def receive_market_data(payload: MarketDataPayload):
    """
    Orquestrador: Receber dados de mercado e gerar sinal com IA

    Fluxo: Market Data → Strategy → Risk → MT5 → Telegram
    """

    # Validar secret
    if payload.secret != WEBHOOK_SECRET:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Secret inválido"
        )

    # Gerar sinal com strategy_service
    strategy_result = generate_signal({
        "pair": payload.pair,
        "timeframe": payload.timeframe,
        "close_prices": payload.close_prices
    })

    # Se sem sinal
    if strategy_result["signal"] is None:
        print(f"⚠️ Sem sinal: {strategy_result['reason']}")
        return {
            "status": "no_signal",
            "reason": strategy_result["reason"],
            "analysis": strategy_result
        }

    # Processar sinal gerado
    signal = process_signal(
        par=payload.pair,
        tipo=strategy_result["signal"],
        timeframe=payload.timeframe,
        preco=strategy_result["indicators"]["price"]
    )

    # Validar risco
    risk_validation = validate_risk(signal)

    # Se bloqueado por risco
    if not risk_validation["approved"]:
        print(f"⚠️ Sinal bloqueado: {risk_validation['reason']}")
        return {
            "status": "blocked",
            "reason": risk_validation["reason"],
            "signal": signal,
            "analysis": strategy_result
        }

    # Conectar ao MT5 e executar ordem
    mt5_result = send_order({
        "pair": signal["par"],
        "signal": signal["tipo"]
    })

    # Se erro no MT5
    if mt5_result["status"] == "erro":
        print(f"❌ Erro MT5: {mt5_result['message']}")
        return {
            "status": "blocked",
            "reason": f"Erro ao executar ordem: {mt5_result['message']}",
            "signal": signal,
            "analysis": strategy_result
        }

    # Montar mensagem Telegram
    message = f"""🚀 <b>{signal['tipo']}</b>
Par: <b>{signal['par']}</b>
Timeframe: <b>{signal['timeframe']}</b>
Preço: <b>{strategy_result['indicators']['price']}</b>
SL: <b>{mt5_result.get('sl', 'N/A')}</b>
TP: <b>{mt5_result.get('tp', 'N/A')}</b>
Confiança: <b>{strategy_result['confidence']}%</b>
Status: <b>EXECUTADO</b>"""

    # Enviar Telegram
    telegram_result = send_telegram_message(message)

    print(f"✅ Ordem executada: {signal['tipo']} {signal['par']} (Conf: {strategy_result['confidence']}%)")

    return {
        "status": "executed",
        "reason": "Ordem executada com sucesso",
        "signal": signal,
        "order": mt5_result,
        "analysis": strategy_result,
        "telegram": telegram_result
    }
