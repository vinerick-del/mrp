"""
Serviço de Estratégia de Trading
Estratégia Agressiva baseada em EMA + RSI
"""

from typing import Dict, Any, List, Optional


def calculate_ema(prices: List[float], period: int) -> Optional[float]:
    """
    Calcular Média Móvel Exponencial (EMA)

    Args:
        prices: Lista de preços
        period: Período da EMA

    Returns:
        Valor da EMA ou None se insuficiente dados
    """
    if len(prices) < period:
        return None

    # Multiplier para EMA
    multiplier = 2 / (period + 1)

    # SMA inicial
    sma = sum(prices[-period:]) / period
    ema = sma

    # Calcular EMA para cada preço após período
    for price in prices[-period + 1:]:
        ema = price * multiplier + ema * (1 - multiplier)

    return round(ema, 5)


def calculate_rsi(prices: List[float], period: int = 14) -> Optional[float]:
    """
    Calcular Índice de Força Relativa (RSI)

    Args:
        prices: Lista de preços
        period: Período do RSI (padrão 14)

    Returns:
        Valor do RSI (0-100) ou None se insuficiente dados
    """
    if len(prices) < period + 1:
        return None

    # Calcular variações
    deltas = [prices[i] - prices[i - 1] for i in range(1, len(prices))]

    # Ganhos e perdas
    gains = [d if d > 0 else 0 for d in deltas]
    losses = [abs(d) if d < 0 else 0 for d in deltas]

    # Médias
    avg_gain = sum(gains[-period:]) / period
    avg_loss = sum(losses[-period:]) / period

    # Evitar divisão por zero
    if avg_loss == 0:
        return 100.0 if avg_gain > 0 else 50.0

    # RS e RSI
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))

    return round(rsi, 2)


def generate_signal(market_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Gerar sinal de trading baseado em EMA + RSI

    Args:
        market_data: Dicionário com pair, timeframe, close_prices

    Returns:
        Dicionário com signal, confidence, reason
    """

    close_prices = market_data.get("close_prices", [])
    pair = market_data.get("pair", "UNKNOWN")
    timeframe = market_data.get("timeframe", "UNKNOWN")

    # Validar dados
    if len(close_prices) < 22:  # Mínimo para EMA21 + RSI
        return {
            "signal": None,
            "confidence": 0,
            "reason": "Dados insuficientes para análise"
        }

    # Calcular indicadores
    ema9 = calculate_ema(close_prices, 9)
    ema21 = calculate_ema(close_prices, 21)
    rsi = calculate_rsi(close_prices, 14)

    # Validar cálculos
    if ema9 is None or ema21 is None or rsi is None:
        return {
            "signal": None,
            "confidence": 0,
            "reason": "Erro ao calcular indicadores"
        }

    current_price = close_prices[-1]

    # Lógica de tendência
    ema_trend = "ALTA" if ema9 > ema21 else "BAIXA"

    # Lógica RSI
    rsi_status = None
    if rsi < 30:
        rsi_status = "SOBREVENDIDO"
    elif rsi > 70:
        rsi_status = "SOBRECOMPRADO"
    else:
        rsi_status = "NEUTRO"

    # Confirmação: tendência + RSI
    signal = None
    confidence = 0
    reason = ""

    # COMPRA: Tendência ALTA + RSI sobrevendido
    if ema_trend == "ALTA" and rsi_status == "SOBREVENDIDO":
        signal = "BUY"
        confidence = 90
        reason = f"EMA9({ema9:.5f}) > EMA21({ema21:.5f}) + RSI({rsi:.2f}) sobrevendido"

    # VENDA: Tendência BAIXA + RSI sobrecomprado
    elif ema_trend == "BAIXA" and rsi_status == "SOBRECOMPRADO":
        signal = "SELL"
        confidence = 90
        reason = f"EMA9({ema9:.5f}) < EMA21({ema21:.5f}) + RSI({rsi:.2f}) sobrecomprado"

    # COMPRA MODERADA: Tendência ALTA (sem RSI negativo)
    elif ema_trend == "ALTA" and rsi_status != "SOBRECOMPRADO":
        signal = "BUY"
        confidence = 60
        reason = f"EMA9({ema9:.5f}) > EMA21({ema21:.5f}) (RSI: {rsi:.2f})"

    # VENDA MODERADA: Tendência BAIXA (sem RSI positivo)
    elif ema_trend == "BAIXA" and rsi_status != "SOBREVENDIDO":
        signal = "SELL"
        confidence = 60
        reason = f"EMA9({ema9:.5f}) < EMA21({ema21:.5f}) (RSI: {rsi:.2f})"

    else:
        signal = None
        confidence = 0
        reason = f"Sem confirmação: {ema_trend} + RSI {rsi:.2f}"

    return {
        "signal": signal,
        "confidence": confidence,
        "reason": reason,
        "pair": pair,
        "timeframe": timeframe,
        "indicators": {
            "ema9": ema9,
            "ema21": ema21,
            "rsi": rsi,
            "price": current_price
        }
    }
