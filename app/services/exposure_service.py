"""
Serviço de Controle de Exposição
Monitoramento de risco total da conta em tempo real
"""

from typing import Dict, Any, List

# Limites de exposição
LOW_EXPOSURE_THRESHOLD = 3.0  # 3%
MEDIUM_EXPOSURE_THRESHOLD = 5.0  # 5%
HIGH_EXPOSURE_THRESHOLD = 10.0  # 10% (crítico)


def check_total_exposure(open_positions: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Verificar exposição total da conta

    Args:
        open_positions: Lista de posições abertas
            [{"symbol": "EURUSD", "lot": 0.02, "risk": 2.0}, ...]

    Returns:
        Dicionário com allowed, risk_level, total_risk
    """
    try:
        # Calcular risco total
        total_risk = sum(pos.get("risk", 0.0) for pos in open_positions)

        # Determinar nível de risco
        if total_risk > MEDIUM_EXPOSURE_THRESHOLD:
            risk_level = "high"
            allowed = False
        elif total_risk >= LOW_EXPOSURE_THRESHOLD:
            risk_level = "medium"
            allowed = True
        else:
            risk_level = "low"
            allowed = True

        return {
            "allowed": allowed,
            "risk_level": risk_level,
            "total_risk": round(total_risk, 2),
            "positions_count": len(open_positions)
        }

    except Exception as e:
        return {
            "allowed": False,
            "risk_level": "error",
            "total_risk": 0.0,
            "message": f"Erro ao calcular exposição: {str(e)}"
        }


def get_open_positions() -> List[Dict[str, Any]]:
    """
    Obter posições abertas (mock simples por enquanto)

    Returns:
        Lista de posições abertas
    """
    try:
        # TODO: Integrar com MT5 via mt5.positions_get()
        # Por enquanto retorna lista vazia (mock)
        positions = []
        return positions

    except Exception as e:
        print(f"Erro ao obter posições abertas: {str(e)}")
        return []


def get_position_risk(lot: float, stop_loss_pips: float, price: float = 1.0) -> float:
    """
    Calcular risco de uma posição em percentual

    Args:
        lot: Tamanho da posição
        stop_loss_pips: Stop loss em pips
        price: Preço atual (padrão 1.0 para simplificar)

    Returns:
        Risco em percentual
    """
    try:
        # Risco = (lot * stop_loss_pips * point) / saldo_conta
        # Simplificado: risco = lot * stop_loss_pips * 0.0001
        risk_percentage = lot * stop_loss_pips * 0.0001 * 100

        return round(risk_percentage, 2)

    except Exception as e:
        print(f"Erro ao calcular risco da posição: {str(e)}")
        return 0.0


def can_open_new_position(
    current_exposure: float,
    new_position_risk: float
) -> Dict[str, Any]:
    """
    Verificar se pode abrir nova posição

    Args:
        current_exposure: Exposição atual total
        new_position_risk: Risco da nova posição

    Returns:
        Dicionário com allowed e motivo
    """
    total_exposure = current_exposure + new_position_risk

    if total_exposure > MEDIUM_EXPOSURE_THRESHOLD:
        return {
            "allowed": False,
            "reason": f"Exposição total ({total_exposure:.2f}%) excede limite ({MEDIUM_EXPOSURE_THRESHOLD}%)",
            "total_exposure": round(total_exposure, 2)
        }

    return {
        "allowed": True,
        "reason": "Exposição dentro dos limites",
        "total_exposure": round(total_exposure, 2)
    }


def get_exposure_status() -> Dict[str, Any]:
    """
    Obter status atual de exposição

    Returns:
        Status detalhado de exposição
    """
    try:
        positions = get_open_positions()
        exposure_check = check_total_exposure(positions)

        return {
            "status": "ok",
            "exposure": exposure_check,
            "thresholds": {
                "low": f"< {LOW_EXPOSURE_THRESHOLD}%",
                "medium": f"{LOW_EXPOSURE_THRESHOLD}% - {MEDIUM_EXPOSURE_THRESHOLD}%",
                "high": f"> {MEDIUM_EXPOSURE_THRESHOLD}%"
            }
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao obter status: {str(e)}"
        }


def reset_positions() -> Dict[str, Any]:
    """
    Resetar rastreamento de posições (útil para testes)

    Returns:
        Confirmação do reset
    """
    return {
        "status": "sucesso",
        "message": "Posições resetadas"
    }
