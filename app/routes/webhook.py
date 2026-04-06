"""
Rotas de Webhook - Orquestrador do Sistema de Trading Multi-Ativo
Fluxo completo: Signal → Risk → Safety → MT5 → Telegram
"""

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from typing import Optional, List
from config import WEBHOOK_SECRET
from services.signal_service import process_signal
from services.strategy_service import generate_signal
from services.risk_management import validate_risk
from services.safety_service import check_safety, update_loss_counter
from services.mt5_service import send_order, connect_mt5
from services.telegram_service import send_telegram_message
from services.performance_service import get_performance_summary

webhook_router = APIRouter(prefix="/webhook", tags=["webhook"])

# Símbolos permitidos para negociação
ALLOWED_SYMBOLS = ["EURUSD", "GBPUSD", "XAUUSD"]


# Modelos de entrada
class SignalPayload(BaseModel):
    secret: str
    symbol: str
    tipo: str
    timeframe: str
    preco: float


class MarketDataPayload(BaseModel):
    secret: str
    symbol: str
    timeframe: str
    close_prices: List[float]


def validate_symbol(symbol: str) -> bool:
    """Validar se símbolo está permitido"""
    return symbol.upper() in ALLOWED_SYMBOLS


@webhook_router.post("")
async def receive_signal(payload: SignalPayload):
    """
    Orquestrador: Receber sinal direto e processar

    Fluxo: Secret → Symbol → Signal → Risk → Safety → MT5 → Telegram
    """

    # Validar secret
    if payload.secret != WEBHOOK_SECRET:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Secret inválido"
        )

    # Validar símbolo
    symbol = payload.symbol.upper()
    if not validate_symbol(symbol):
        print(f"⚠️ Símbolo não permitido: {symbol}")
        return {
            "status": "rejected",
            "reason": f"Símbolo {symbol} não está na lista permitida",
            "allowed_symbols": ALLOWED_SYMBOLS
        }

    print(f"🔄 Processando ativo: {symbol}")

    # Processar sinal
    signal = process_signal(
        par=symbol,
        tipo=payload.tipo,
        timeframe=payload.timeframe,
        preco=payload.preco
    )

    # Validar risco
    risk_validation = validate_risk(signal)

    # Se bloqueado por risco
    if not risk_validation["approved"]:
        print(f"⚠️ Sinal bloqueado por risco: {risk_validation['reason']}")
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": risk_validation["reason"],
            "signal": signal
        }

    # Validar safety
    performance = get_performance_summary()
    safety_check = check_safety(performance)

    # Se bloqueado por safety
    if not safety_check["allowed"]:
        print(f"⛔ {safety_check['reason']}")
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": safety_check["reason"],
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
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": f"Erro ao executar ordem: {mt5_result['message']}",
            "signal": signal
        }

    # Montar mensagem Telegram
    message = f"""🚀 <b>{signal['tipo']}</b>
Par: <b>{symbol}</b>
Timeframe: <b>{signal['timeframe']}</b>
Preço: <b>{signal['preco']}</b>
SL: <b>{mt5_result.get('sl', 'N/A')}</b>
TP: <b>{mt5_result.get('tp', 'N/A')}</b>
Volume: <b>{mt5_result.get('volume', 'N/A')}</b>
Status: <b>EXECUTADO</b>"""

    # Enviar Telegram
    telegram_result = send_telegram_message(message)

    print(f"✅ Ordem executada: {signal['tipo']} {symbol}")

    return {
        "status": "executed",
        "reason": "Ordem executada com sucesso",
        "symbol": symbol,
        "signal": signal,
        "order": mt5_result,
        "telegram": telegram_result
    }


@webhook_router.post("/market-data")
async def receive_market_data(payload: MarketDataPayload):
    """
    Orquestrador: Receber dados de mercado e gerar sinal com IA

    Fluxo: Market Data → Strategy → Risk → Safety → MT5 → Telegram
    """

    # Validar secret
    if payload.secret != WEBHOOK_SECRET:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Secret inválido"
        )

    # Validar símbolo
    symbol = payload.symbol.upper()
    if not validate_symbol(symbol):
        print(f"⚠️ Símbolo não permitido: {symbol}")
        return {
            "status": "rejected",
            "reason": f"Símbolo {symbol} não está na lista permitida",
            "allowed_symbols": ALLOWED_SYMBOLS
        }

    print(f"🔄 Processando ativo: {symbol}")

    # Gerar sinal com strategy_service
    strategy_result = generate_signal({
        "pair": symbol,
        "timeframe": payload.timeframe,
        "close_prices": payload.close_prices
    })

    # Se sem sinal
    if strategy_result["signal"] is None:
        print(f"⚠️ Sem sinal para {symbol}: {strategy_result['reason']}")
        return {
            "status": "no_signal",
            "reason": strategy_result["reason"],
            "symbol": symbol,
            "analysis": strategy_result
        }

    # Processar sinal gerado
    signal = process_signal(
        par=symbol,
        tipo=strategy_result["signal"],
        timeframe=payload.timeframe,
        preco=strategy_result["indicators"]["price"]
    )

    # Validar risco
    risk_validation = validate_risk(signal)

    # Se bloqueado por risco
    if not risk_validation["approved"]:
        print(f"⚠️ Sinal bloqueado por risco: {risk_validation['reason']}")
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": risk_validation["reason"],
            "symbol": symbol,
            "signal": signal,
            "analysis": strategy_result
        }

    # Validar safety
    performance = get_performance_summary()
    safety_check = check_safety(performance)

    # Se bloqueado por safety
    if not safety_check["allowed"]:
        print(f"⛔ {safety_check['reason']}")
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": safety_check["reason"],
            "symbol": symbol,
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
        update_loss_counter(True)
        return {
            "status": "blocked",
            "reason": f"Erro ao executar ordem: {mt5_result['message']}",
            "symbol": symbol,
            "signal": signal,
            "analysis": strategy_result
        }

    # Montar mensagem Telegram
    message = f"""🚀 <b>{signal['tipo']}</b>
Par: <b>{symbol}</b>
Timeframe: <b>{signal['timeframe']}</b>
Preço: <b>{strategy_result['indicators']['price']}</b>
SL: <b>{mt5_result.get('sl', 'N/A')}</b>
TP: <b>{mt5_result.get('tp', 'N/A')}</b>
Volume: <b>{mt5_result.get('volume', 'N/A')}</b>
Confiança: <b>{strategy_result['confidence']}%</b>
Status: <b>EXECUTADO</b>"""

    # Enviar Telegram
    telegram_result = send_telegram_message(message)

    print(f"✅ Ordem executada: {signal['tipo']} {symbol} (Conf: {strategy_result['confidence']}%)")

    return {
        "status": "executed",
        "reason": "Ordem executada com sucesso",
        "symbol": symbol,
        "signal": signal,
        "order": mt5_result,
        "analysis": strategy_result,
        "telegram": telegram_result
    }
