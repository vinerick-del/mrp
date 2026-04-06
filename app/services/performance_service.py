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


def get_symbol_ranking() -> List[Dict[str, Any]]:
    """
    Obter ranking de ativos baseado em performance

    Calcula score = (win_rate * 0.6) + (profit * 0.4)
    Ordena do melhor para o pior

    Returns:
        Lista de ativos ordenada por performance
    """
    if len(trades_history) == 0:
        return []

    # Agrupar trades por símbolo
    symbols_data = {}

    for trade in trades_history:
        symbol = trade["par"]
        if symbol not in symbols_data:
            symbols_data[symbol] = {
                "trades": [],
                "wins": 0,
                "losses": 0,
                "profit": 0.0
            }

        symbols_data[symbol]["trades"].append(trade)
        if trade["result"] == "WIN":
            symbols_data[symbol]["wins"] += 1
        elif trade["result"] == "LOSS":
            symbols_data[symbol]["losses"] += 1

        symbols_data[symbol]["profit"] += trade["profit_loss"]

    # Calcular ranking
    ranking = []

    for symbol, data in symbols_data.items():
        total_trades = len(data["trades"])
        if total_trades == 0:
            continue

        win_rate = (data["wins"] / total_trades * 100)
        profit = data["profit"]

        # Score normalizado
        # Win rate: 0-100, normalizamos para 0-1
        # Profit: pode ser negativo, normalizamos pela magnitude
        win_rate_normalized = win_rate / 100.0
        profit_normalized = min(profit / 1000.0, 1.0) if profit >= 0 else max(profit / 1000.0, -1.0)

        score = (win_rate_normalized * 0.6) + (profit_normalized * 0.4)
        score = max(0, min(1.0, score))  # Limitar entre 0 e 1

        ranking.append({
            "symbol": symbol,
            "win_rate": round(win_rate, 2),
            "profit": round(profit, 2),
            "trades": total_trades,
            "wins": data["wins"],
            "losses": data["losses"],
            "score": round(score, 2)
        })

    # Ordenar do melhor para o pior
    ranking.sort(key=lambda x: x["score"], reverse=True)

    return ranking


def get_best_symbol() -> Optional[Dict[str, Any]]:
    """
    Obter ativo com melhor performance

    Returns:
        Dicionário com melhor ativo ou None
    """
    ranking = get_symbol_ranking()
    return ranking[0] if ranking else None


def get_worst_symbol() -> Optional[Dict[str, Any]]:
    """
    Obter ativo com pior performance

    Returns:
        Dicionário com pior ativo ou None
    """
    ranking = get_symbol_ranking()
    return ranking[-1] if ranking else None
