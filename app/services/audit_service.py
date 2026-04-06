"""
Serviço de Auditoria
Log estruturado de todas as decisões do sistema
"""

from datetime import datetime
from typing import Dict, Any, List, Optional

# Armazenamento em memória (será expandido para database)
audit_log: List[Dict[str, Any]] = []

# Tipos de eventos válidos
VALID_EVENT_TYPES = {
    "signal_received",
    "signal_processed",
    "risk_blocked",
    "safety_blocked",
    "exposure_blocked",
    "order_executed",
    "order_failed"
}


def log_event(event_type: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Registrar evento no log de auditoria

    Args:
        event_type: Tipo de evento (signal_received, signal_processed, etc)
        data: Dicionário com detalhes do evento
            - symbol: Par de moedas
            - details: Informações adicionais

    Returns:
        Confirmação do registro
    """
    try:
        # Validar tipo de evento
        if event_type not in VALID_EVENT_TYPES:
            return {
                "status": "erro",
                "message": f"Tipo de evento inválido: {event_type}",
                "valid_types": list(VALID_EVENT_TYPES)
            }

        # Criar entrada de log
        log_entry = {
            "timestamp": datetime.utcnow().isoformat(),
            "event_type": event_type,
            "symbol": data.get("symbol", "UNKNOWN"),
            "details": data.get("details", {})
        }

        audit_log.append(log_entry)

        return {
            "status": "sucesso",
            "message": f"Evento '{event_type}' registrado",
            "total_logs": len(audit_log)
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao registrar evento: {str(e)}"
        }


def get_logs(
    event_type: Optional[str] = None,
    symbol: Optional[str] = None,
    limit: int = 100
) -> List[Dict[str, Any]]:
    """
    Obter logs com filtros opcionais

    Args:
        event_type: Filtrar por tipo de evento (opcional)
        symbol: Filtrar por símbolo (opcional)
        limit: Limite de resultados

    Returns:
        Lista de logs filtrada
    """
    try:
        filtered_logs = audit_log

        # Filtrar por tipo de evento
        if event_type:
            filtered_logs = [
                log for log in filtered_logs
                if log["event_type"] == event_type
            ]

        # Filtrar por símbolo
        if symbol:
            filtered_logs = [
                log for log in filtered_logs
                if log["symbol"].upper() == symbol.upper()
            ]

        # Aplicar limite (retornar os últimos N logs)
        return filtered_logs[-limit:]

    except Exception as e:
        print(f"Erro ao obter logs: {str(e)}")
        return []


def get_logs_by_type(event_type: str) -> List[Dict[str, Any]]:
    """
    Obter todos os logs de um tipo específico

    Args:
        event_type: Tipo de evento

    Returns:
        Lista de logs do tipo
    """
    return get_logs(event_type=event_type)


def get_logs_by_symbol(symbol: str) -> List[Dict[str, Any]]:
    """
    Obter todos os logs de um símbolo específico

    Args:
        symbol: Par de moedas

    Returns:
        Lista de logs do símbolo
    """
    return get_logs(symbol=symbol)


def get_event_summary() -> Dict[str, Any]:
    """
    Obter resumo dos eventos registrados

    Returns:
        Estatísticas de eventos
    """
    if len(audit_log) == 0:
        return {
            "total_events": 0,
            "by_type": {},
            "by_symbol": {}
        }

    # Contar eventos por tipo
    by_type = {}
    for event_type in VALID_EVENT_TYPES:
        count = sum(1 for log in audit_log if log["event_type"] == event_type)
        if count > 0:
            by_type[event_type] = count

    # Contar eventos por símbolo
    by_symbol = {}
    for log in audit_log:
        symbol = log["symbol"]
        by_symbol[symbol] = by_symbol.get(symbol, 0) + 1

    return {
        "total_events": len(audit_log),
        "by_type": by_type,
        "by_symbol": by_symbol,
        "event_types_available": list(VALID_EVENT_TYPES)
    }


def get_recent_events(minutes: int = 60) -> List[Dict[str, Any]]:
    """
    Obter eventos dos últimos N minutos

    Args:
        minutes: Número de minutos

    Returns:
        Lista de eventos recentes
    """
    from datetime import timedelta

    cutoff_time = datetime.utcnow() - timedelta(minutes=minutes)

    return [
        log for log in audit_log
        if datetime.fromisoformat(log["timestamp"]) > cutoff_time
    ]


def get_failed_orders() -> List[Dict[str, Any]]:
    """
    Obter todas as ordens que falharam

    Returns:
        Lista de ordens falhadas
    """
    return get_logs_by_type("order_failed")


def get_blocked_signals() -> List[Dict[str, Any]]:
    """
    Obter todos os sinais bloqueados

    Returns:
        Lista de sinais bloqueados
    """
    blocked = []
    blocked.extend(get_logs_by_type("risk_blocked"))
    blocked.extend(get_logs_by_type("safety_blocked"))
    blocked.extend(get_logs_by_type("exposure_blocked"))

    return sorted(blocked, key=lambda x: x["timestamp"], reverse=True)


def clear_logs() -> Dict[str, Any]:
    """
    Limpar todos os logs (útil para testes)

    Returns:
        Confirmação da limpeza
    """
    global audit_log
    audit_log = []

    return {
        "status": "sucesso",
        "message": "Logs limpos"
    }


def export_logs(event_type: Optional[str] = None) -> Dict[str, Any]:
    """
    Exportar logs para análise

    Args:
        event_type: Tipo de evento a exportar (opcional)

    Returns:
        Dicionário com logs e metadados
    """
    logs = get_logs(event_type=event_type) if event_type else audit_log

    return {
        "status": "sucesso",
        "total_logs": len(logs),
        "event_type_filter": event_type,
        "export_timestamp": datetime.utcnow().isoformat(),
        "logs": logs
    }
