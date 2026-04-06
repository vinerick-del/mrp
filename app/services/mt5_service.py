"""
Serviço de Integração com MetaTrader 5
Execução automática de ordens com cálculo de risco
"""

from typing import Dict, Any, Optional
import MetaTrader5 as mt5
from config import RISK_PER_TRADE

# Configurações de ordem
DEFAULT_VOLUME = 0.01  # Lote padrão (mínimo)
SL_PIPS = 20  # Stop Loss em pips
TP_PIPS = 40  # Take Profit em pips (RR 1:2)


def connect_mt5(server: str = "MetaQuotes-Demo", login: int = None, password: str = None) -> Dict[str, Any]:
    """
    Conectar ao MetaTrader 5

    Args:
        server: Nome do servidor MT5
        login: Login da conta (opcional)
        password: Senha da conta (opcional)

    Returns:
        Status de conexão
    """
    try:
        if not mt5.initialize():
            return {
                "status": "erro",
                "message": f"Falha ao inicializar MT5: {mt5.last_error()}",
                "connected": False
            }

        # Se login e password fornecidos, conectar à conta específica
        if login and password:
            if not mt5.login(login, password, server):
                return {
                    "status": "erro",
                    "message": f"Falha ao login: {mt5.last_error()}",
                    "connected": False
                }

        return {
            "status": "sucesso",
            "message": "Conectado ao MetaTrader 5",
            "connected": True
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro de conexão: {str(e)}",
            "connected": False
        }


def disconnect_mt5() -> Dict[str, Any]:
    """Desconectar do MetaTrader 5"""
    try:
        mt5.shutdown()
        return {
            "status": "sucesso",
            "message": "Desconectado do MetaTrader 5"
        }
    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao desconectar: {str(e)}"
        }


def get_symbol_info(symbol: str) -> Optional[Dict[str, Any]]:
    """
    Obter informações do símbolo (preço, pips)

    Args:
        symbol: Par de moedas (ex: EURUSD)

    Returns:
        Dicionário com preços ou None se erro
    """
    try:
        symbol_info = mt5.symbol_info(symbol)
        if symbol_info is None:
            return None

        return {
            "symbol": symbol,
            "bid": symbol_info.bid,
            "ask": symbol_info.ask,
            "point": symbol_info.point
        }

    except Exception as e:
        print(f"Erro ao obter info: {str(e)}")
        return None


def has_open_position(symbol: str) -> bool:
    """
    Verificar se existe posição aberta para o símbolo

    Args:
        symbol: Par de moedas (ex: EURUSD)

    Returns:
        True se existe posição aberta, False caso contrário
    """
    try:
        positions = mt5.positions_get(symbol=symbol)
        if positions is None:
            return False
        return len(positions) > 0
    except Exception as e:
        print(f"Erro ao verificar posições: {str(e)}")
        return False


def calculate_sl_tp(signal_type: str, current_price: float, point: float) -> Dict[str, float]:
    """
    Calcular Stop Loss e Take Profit

    Args:
        signal_type: "BUY" ou "SELL"
        current_price: Preço atual
        point: Um pip da moeda

    Returns:
        Dicionário com SL e TP
    """
    pip_value = point * SL_PIPS

    if signal_type == "BUY":
        sl = current_price - pip_value
        tp = current_price + (pip_value * 2)  # RR 1:2
    else:  # SELL
        sl = current_price + pip_value
        tp = current_price - (pip_value * 2)  # RR 1:2

    return {
        "sl": round(sl, 5),
        "tp": round(tp, 5)
    }


def calculate_position_size(sl_pips: float, symbol: str) -> Optional[float]:
    """
    Calcular tamanho da posição baseado no risco por trade

    Args:
        sl_pips: Stop Loss em pips
        symbol: Par de moedas

    Returns:
        Tamanho do lote calculado ou None se erro
    """
    try:
        # Obter saldo da conta
        account_info = mt5.account_info()
        if account_info is None:
            print("Erro ao obter informações da conta")
            return None

        balance = account_info.balance

        # Calcular risco em reais
        risk_amount = balance * (RISK_PER_TRADE / 100)

        # Obter informações do símbolo
        symbol_info = get_symbol_info(symbol)
        if symbol_info is None:
            return None

        # Valor por pip
        point = symbol_info["point"]
        bid = symbol_info["bid"]

        # Para Forex, valor por pip = point * 10000
        # Ajustar conforme necessário para outros ativos
        value_per_pip = point * 10000 * bid

        # Calcular lote
        if value_per_pip <= 0:
            return DEFAULT_VOLUME

        lot = risk_amount / (sl_pips * value_per_pip)

        # Respeitar lote mínimo
        lot = max(lot, DEFAULT_VOLUME)

        # Arredondar para múltiplos de 0.01
        lot = round(lot, 2)

        return lot

    except Exception as e:
        print(f"Erro ao calcular tamanho da posição: {str(e)}")
        return DEFAULT_VOLUME


def send_order(signal: Dict[str, Any]) -> Dict[str, Any]:
    """
    Enviar ordem ao MetaTrader 5

    Args:
        signal: Dicionário com pair, signal (BUY/SELL)

    Returns:
        Status da operação
    """
    try:
        symbol = signal.get("pair")
        signal_type = signal.get("signal")

        if not symbol or not signal_type:
            return {
                "status": "erro",
                "message": "Sinal inválido (faltam pair ou signal)"
            }

        # Verificar se já existe posição aberta
        if has_open_position(symbol):
            return {
                "status": "erro",
                "message": f"Já existe posição aberta para {symbol}"
            }

        # Obter informações do símbolo
        symbol_info = get_symbol_info(symbol)
        if symbol_info is None:
            return {
                "status": "erro",
                "message": f"Símbolo {symbol} não encontrado"
            }

        # Preço para operação
        price = symbol_info["ask"] if signal_type == "BUY" else symbol_info["bid"]

        # Calcular SL e TP
        sl_tp = calculate_sl_tp(signal_type, price, symbol_info["point"])

        # Calcular tamanho da posição baseado em risco
        position_size = calculate_position_size(SL_PIPS, symbol)
        if position_size is None:
            position_size = DEFAULT_VOLUME

        # Tipo de ordem
        order_type = mt5.ORDER_TYPE_BUY if signal_type == "BUY" else mt5.ORDER_TYPE_SELL

        # Montar requisição de ordem
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": symbol,
            "volume": position_size,
            "type": order_type,
            "price": price,
            "sl": sl_tp["sl"],
            "tp": sl_tp["tp"],
            "deviation": 10,
            "magic": 234000,
            "comment": f"Trading System - {signal_type}",
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_IOC
        }

        # Enviar ordem
        result = mt5.order_send(request)

        if result.retcode != mt5.TRADE_RETCODE_DONE:
            return {
                "status": "erro",
                "message": f"Falha ao enviar ordem: {result.comment}",
                "retcode": result.retcode
            }

        return {
            "status": "sucesso",
            "message": f"Ordem {signal_type} enviada",
            "order_id": result.order,
            "price": price,
            "sl": sl_tp["sl"],
            "tp": sl_tp["tp"],
            "volume": position_size
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao enviar ordem: {str(e)}"
        }
