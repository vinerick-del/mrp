"""
Serviço de Análise de Performance
Registro e análise de operações executadas
"""

from datetime import datetime
from typing import Dict, Any, List, Optional

# Armazenamento em memória (será expandido para database)
trades_history: List[Dict[str, Any]] = []


def log_trade(trade_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Registrar uma operação executada

    Args:
        trade_data: Dicionário com dados da operação
            - par: Par de moedas
            - tipo: BUY ou SELL
            - volume: Tamanho da posição (lote)
            - entry_price: Preço de entrada
            - exit_price: Preço de saída
            - profit_loss: Lucro ou prejuízo
            - result: WIN ou LOSS

    Returns:
        Confirmação do registro
    """
    try:
        trade = {
            "par": trade_data.get("par"),
            "tipo": trade_data.get("tipo"),
            "volume": trade_data.get("volume"),
            "entry_price": trade_data.get("entry_price"),
            "exit_price": trade_data.get("exit_price"),
            "profit_loss": float(trade_data.get("profit_loss", 0)),
            "result": trade_data.get("result"),  # WIN ou LOSS
            "timestamp": datetime.utcnow().isoformat(),
            "duration": trade_data.get("duration")  # Tempo da operação em minutos
        }

        trades_history.append(trade)

        return {
            "status": "sucesso",
            "message": "Trade registrado",
            "total_trades": len(trades_history)
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao registrar trade: {str(e)}"
        }


def get_performance_summary() -> Dict[str, Any]:
    """
    Obter resumo de performance

    Returns:
        Dicionário com estatísticas de performance
    """
    if len(trades_history) == 0:
        return {
            "total_trades": 0,
            "wins": 0,
            "losses": 0,
            "win_rate": 0.0,
            "total_profit": 0.0,
            "total_loss": 0.0,
            "net_profit": 0.0,
            "average_win": 0.0,
            "average_loss": 0.0,
            "profit_factor": 0.0,
            "drawdown": 0.0,
            "largest_win": 0.0,
            "largest_loss": 0.0
        }

    # Calcular estatísticas
    total_trades = len(trades_history)
    wins = sum(1 for t in trades_history if t["result"] == "WIN")
    losses = sum(1 for t in trades_history if t["result"] == "LOSS")

    win_rate = (wins / total_trades * 100) if total_trades > 0 else 0.0

    # Lucros e prejuízos
    total_profit = sum(t["profit_loss"] for t in trades_history if t["profit_loss"] > 0)
    total_loss = abs(sum(t["profit_loss"] for t in trades_history if t["profit_loss"] < 0))
    net_profit = sum(t["profit_loss"] for t in trades_history)

    # Média de ganho e perda
    winning_trades = [t["profit_loss"] for t in trades_history if t["profit_loss"] > 0]
    losing_trades = [t["profit_loss"] for t in trades_history if t["profit_loss"] < 0]

    average_win = (sum(winning_trades) / len(winning_trades)) if winning_trades else 0.0
    average_loss = (abs(sum(losing_trades)) / len(losing_trades)) if losing_trades else 0.0

    # Profit Factor
    profit_factor = (total_profit / total_loss) if total_loss > 0 else 0.0

    # Drawdown (maior perda acumulada)
    drawdown = calculate_drawdown()

    # Maior ganho e perda
    largest_win = max(winning_trades) if winning_trades else 0.0
    largest_loss = abs(min(losing_trades)) if losing_trades else 0.0

    return {
        "total_trades": total_trades,
        "wins": wins,
        "losses": losses,
        "win_rate": round(win_rate, 2),
        "total_profit": round(total_profit, 2),
        "total_loss": round(total_loss, 2),
        "net_profit": round(net_profit, 2),
        "average_win": round(average_win, 2),
        "average_loss": round(average_loss, 2),
        "profit_factor": round(profit_factor, 2),
        "drawdown": round(drawdown, 2),
        "largest_win": round(largest_win, 2),
        "largest_loss": round(largest_loss, 2)
    }


def calculate_drawdown() -> float:
    """
    Calcular maior perda acumulada (drawdown)

    Returns:
        Valor do drawdown
    """
    if len(trades_history) == 0:
        return 0.0

    cumulative_profit = 0.0
    peak = 0.0
    max_drawdown = 0.0

    for trade in trades_history:
        cumulative_profit += trade["profit_loss"]
        if cumulative_profit > peak:
            peak = cumulative_profit

        drawdown = peak - cumulative_profit
        if drawdown > max_drawdown:
            max_drawdown = drawdown

    return max_drawdown


def get_trades_by_pair(pair: str) -> List[Dict[str, Any]]:
    """
    Obter histórico de trades por par

    Args:
        pair: Par de moedas

    Returns:
        Lista de trades do par
    """
    return [t for t in trades_history if t["par"] == pair]


def get_trades_by_timeframe(hours: int = 24) -> List[Dict[str, Any]]:
    """
    Obter trades dos últimas N horas

    Args:
        hours: Número de horas

    Returns:
        Lista de trades recentes
    """
    from datetime import timedelta

    cutoff_time = datetime.utcnow() - timedelta(hours=hours)

    return [
        t for t in trades_history
        if datetime.fromisoformat(t["timestamp"]) > cutoff_time
    ]


def reset_history() -> Dict[str, Any]:
    """Resetar histórico de trades (útil para testes)"""
    global trades_history
    trades_history = []

    return {
        "status": "sucesso",
        "message": "Histórico resetado"
    }
