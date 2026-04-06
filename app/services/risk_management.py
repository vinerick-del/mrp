"""
Serviço de Gerenciamento de Risco
"""

from datetime import datetime, timedelta
from typing import Dict, Any, List
from config import MAX_TRADES_PER_DAY, RISK_PER_TRADE

# Estado em memória (será expandido para database)
trades_today: List[Dict[str, Any]] = []
recent_signals: List[Dict[str, Any]] = []


def is_within_trading_hours() -> bool:
    """
    Validar se está dentro do horário de negociação (08:00 - 18:00)

    Returns:
        True se dentro do horário, False caso contrário
    """
    now = datetime.utcnow()
    hour = now.hour
    return 8 <= hour < 18


def check_daily_limit() -> bool:
    """
    Verificar se não excedeu limite de trades por dia

    Returns:
        True se dentro do limite, False se excedeu
    """
    today = datetime.utcnow().date()

    # Contar trades de hoje
    trades_count = sum(
        1 for trade in trades_today
        if datetime.fromisoformat(trade["timestamp"]).date() == today
    )

    return trades_count < MAX_TRADES_PER_DAY


def check_duplicate_signal(signal: Dict[str, Any]) -> bool:
    """
    Verificar se há sinal duplicado recente (mesmo PAR + TIMEFRAME + TIPO)

    Args:
        signal: Sinal a verificar

    Returns:
        True se não há duplicado, False se há
    """
    now = datetime.utcnow()
    threshold = now - timedelta(minutes=5)  # Janela de 5 minutos

    # Remover sinais antigos
    global recent_signals
    recent_signals = [
        s for s in recent_signals
        if datetime.fromisoformat(s["timestamp"]) > threshold
    ]

    # Verificar duplicatas
    for recent in recent_signals:
        if (
            recent["par"] == signal["par"]
            and recent["timeframe"] == signal["timeframe"]
            and recent["tipo"] == signal["tipo"]
        ):
            return False  # Encontrou duplicata

    return True  # Sem duplicata


def validate_risk(signal: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validar risco do sinal antes de autorizar

    Args:
        signal: Sinal com par, tipo, timeframe, preco, timestamp

    Returns:
        Dicionário com approved (bool) e reason (str)
    """

    # Validar horário de negociação
    if not is_within_trading_hours():
        return {
            "approved": False,
            "reason": "Fora do horário de negociação (08:00 - 18:00)"
        }

    # Validar limite diário
    if not check_daily_limit():
        return {
            "approved": False,
            "reason": f"Limite de {MAX_TRADES_PER_DAY} trades por dia atingido"
        }

    # Validar sinal duplicado
    if not check_duplicate_signal(signal):
        return {
            "approved": False,
            "reason": "Sinal duplicado detectado (mesmo par, timeframe, tipo)"
        }

    # Preparado para RISK_PER_TRADE
    # TODO: Implementar validação de risco monetário

    # Sinal aprovado
    # Registrar sinal
    recent_signals.append(signal)
    trades_today.append(signal)

    return {
        "approved": True,
        "reason": "Sinal aprovado pelo gerenciador de risco"
    }


def reset_daily_trades():
    """Resetar contador de trades diários (chamar em início do dia)"""
    global trades_today
    today = datetime.utcnow().date()
    trades_today = [
        t for t in trades_today
        if datetime.fromisoformat(t["timestamp"]).date() == today
    ]
