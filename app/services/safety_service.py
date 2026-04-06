"""
Serviço de Proteção de Capital
Kill Switch automático baseado em performance
"""

from typing import Dict, Any

# Estado do safety manager
consecutive_losses = 0
safety_active = True


def check_safety(performance: Dict[str, Any]) -> Dict[str, Any]:
    """
    Verificar se operações são permitidas baseado em proteções

    Args:
        performance: Dicionário com estatísticas de performance

    Returns:
        Dicionário com allowed (bool) e reason (str)
    """
    global consecutive_losses

    # Se safety está desativado, permitir
    if not safety_active:
        return {
            "allowed": True,
            "reason": "Safety temporariamente desativado"
        }

    # Extrair métricas
    drawdown = performance.get("drawdown", 0)
    win_rate = performance.get("win_rate", 0)
    total_trades = performance.get("total_trades", 0)
    wins = performance.get("wins", 0)
    losses = performance.get("losses", 0)

    # REGRA 1: Drawdown crítico (>= 10%)
    if drawdown >= 10.0:
        return {
            "allowed": False,
            "reason": f"⛔ Kill Switch ativado: Drawdown crítico ({drawdown:.2f}%)"
        }

    # REGRA 2: 3 losses consecutivos
    if consecutive_losses >= 3:
        return {
            "allowed": False,
            "reason": f"⛔ Kill Switch ativado: {consecutive_losses} perdas consecutivas"
        }

    # REGRA 3: Win rate baixo após 20 trades
    if total_trades >= 20 and win_rate < 40.0:
        return {
            "allowed": False,
            "reason": f"⛔ Kill Switch ativado: Win rate baixo ({win_rate:.2f}%) após {total_trades} trades"
        }

    # Todas as verificações passaram
    return {
        "allowed": True,
        "reason": "Operações permitidas"
    }


def update_loss_counter(is_loss: bool) -> None:
    """
    Atualizar contador de losses consecutivos

    Args:
        is_loss: True se a operação foi perda, False se foi ganho
    """
    global consecutive_losses

    if is_loss:
        consecutive_losses += 1
    else:
        consecutive_losses = 0


def reset_safety() -> Dict[str, Any]:
    """
    Resetar contadores de safety (útil para testes ou reset manual)

    Returns:
        Confirmação do reset
    """
    global consecutive_losses

    consecutive_losses = 0

    return {
        "status": "sucesso",
        "message": "Safety resetado",
        "consecutive_losses": consecutive_losses
    }


def deactivate_safety() -> Dict[str, Any]:
    """
    Desativar proteção de capital (temporariamente)

    Returns:
        Confirmação da desativação
    """
    global safety_active

    safety_active = False

    return {
        "status": "sucesso",
        "message": "Safety desativado temporariamente"
    }


def activate_safety() -> Dict[str, Any]:
    """
    Ativar proteção de capital

    Returns:
        Confirmação da ativação
    """
    global safety_active

    safety_active = True

    return {
        "status": "sucesso",
        "message": "Safety ativado"
    }


def get_safety_status() -> Dict[str, Any]:
    """
    Obter status atual do safety manager

    Returns:
        Status de proteção
    """
    return {
        "safety_active": safety_active,
        "consecutive_losses": consecutive_losses,
        "status": "active" if safety_active else "inactive"
    }
