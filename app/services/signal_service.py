"""
Serviço de Processamento de Sinais de Trading
"""

from datetime import datetime
from typing import Dict, Any


def process_signal(par: str, tipo: str, timeframe: str, preco: float) -> Dict[str, Any]:
    """
    Processar sinal recebido do webhook

    Args:
        par: Par de moedas (ex: EURUSD)
        tipo: Tipo de operação (ex: COMPRA, VENDA)
        timeframe: Timeframe (ex: 5M, 1H)
        preco: Preço de entrada

    Returns:
        Dicionário com sinal estruturado e timestamp
    """

    signal = {
        "par": par,
        "tipo": tipo,
        "timeframe": timeframe,
        "preco": preco,
        "timestamp": datetime.utcnow().isoformat()
    }

    return signal


def validate_signal_data(signal: Dict[str, Any]) -> bool:
    """
    Validar estrutura do sinal processado

    Args:
        signal: Sinal processado

    Returns:
        True se válido, False caso contrário
    """

    required_fields = ["par", "tipo", "timeframe", "preco", "timestamp"]
    return all(field in signal for field in required_fields)
